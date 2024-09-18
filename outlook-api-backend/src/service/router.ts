import { errorHandler } from '@backstage/backend-common';
import { LoggerService } from '@backstage/backend-plugin-api';
import express from 'express';
import Router from 'express-promise-router';
import { Config } from '@backstage/config';
import {
  ExchangeService,
  ExchangeVersion,
  WebCredentials,
  Uri,
  WellKnownFolderName,
  ItemView,
  PropertySet,
  BasePropertySet,
  EmailMessage,
  EmailMessageSchema,
  DateTime,
  Appointment,
  AppointmentSchema,
  CalendarView,
  FolderId,
  Folder,
  FolderView,
  FindFoldersResults,
  FolderSchema,
  FolderTraversal,
  SearchFilter,
} from 'ews-javascript-api';
import Redis from 'ioredis';
import crypto from 'crypto';
import cookieSession from 'cookie-session';

export interface RouterOptions {
  logger: LoggerService;
  config: Config;
}

export interface EmailItem {
  subject: string;
  receivedDate: string;
  from: string;
  link: string;
}

interface CalendarEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  location?: string;
  isAllDay: boolean;
  isRecurring: boolean;
}

export async function createRouter(
  options: RouterOptions,
): Promise<express.Router> {
  const { logger, config } = options;

  const router = Router();
  router.use(express.json());

  // Redis client setup
  const redisClient = new Redis({
    host: config.getString('exampleOutlook.redis.host'),
    port: config.getNumber('exampleOutlook.redis.port'),
  });
  redisClient.on('error', (err) => console.log('Redis Client Error', err));

  // Cookie session middleware
  router.use(cookieSession({
    name: 'session',
    keys: [config.getString('exampleOutlook.sessionSecret')],
    maxAge: 365 * 24 * 60 * 60 * 1000 // 1 year
  }));

  // AES encryption setup
  const aesKeyHex = config.getString('exampleOutlook.aesSecretKey');
  const aesKey = Buffer.from(aesKeyHex, 'hex');

  // Encryption helper functions
  const encrypt = (text: string) => {
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv('aes-256-cbc', Buffer.from(aesKey), iv);
    let encrypted = cipher.update(text);
    encrypted = Buffer.concat([encrypted, cipher.final()]);
    return iv.toString('hex') + ':' + encrypted.toString('hex');
  };

  const decrypt = (text: string) => {
    const textParts = text.split(':');
    const iv = Buffer.from(textParts.shift()!, 'hex');
    const encryptedText = Buffer.from(textParts.join(':'), 'hex');
    const decipher = crypto.createDecipheriv('aes-256-cbc', Buffer.from(aesKey), iv);
    let decrypted = decipher.update(encryptedText);
    decrypted = Buffer.concat([decrypted, decipher.final()]);
    return decrypted.toString();
  };

  // Auth middleware
  interface CustomRequest extends express.Request {
    userCredentials?: any;
  }
  
  // @ts-ignore
  const authMiddleware = async (req: CustomRequest, res: express.Response, next: express.NextFunction) => {
    const userEmail = req.session?.userEmail;
    if (!userEmail) {
      return res.status(401).json({ error: 'Not logged in' });
    }
    
    const encryptedCredentials = await redisClient.hget(userEmail, 'password');
    if (!encryptedCredentials) {
      return res.status(401).json({ error: 'Invalid session' });
    }
    
    const credentials = JSON.parse(decrypt(encryptedCredentials));
    req.userCredentials = credentials;
    next();
  };

  router.get('/health', (_, response) => {
    logger.info('PONG!');
    response.json({ status: 'ok' });
  });
  
  async function listCalendars(service: ExchangeService, rootFolderId: WellKnownFolderName | FolderId) {
    const folderView = new FolderView(100);
    folderView.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.FolderClass);
    folderView.Traversal = FolderTraversal.Deep;
  
    const searchFilter = new SearchFilter.IsEqualTo(FolderSchema.FolderClass, "IPF.Appointment");
  
    try {
      let findFoldersResults: FindFoldersResults;
  
      if (rootFolderId instanceof FolderId) {
        findFoldersResults = await service.FindFolders(rootFolderId, searchFilter, folderView);
      } else {
        findFoldersResults = await service.FindFolders(rootFolderId, searchFilter, folderView);
      }
  
      return findFoldersResults.Folders
        .filter(folder => folder.FolderClass === "IPF.Appointment")
        .map(folder => ({
          id: folder.Id.UniqueId,
          name: folder.DisplayName
        }));
    } catch (error) {
      console.error('Error listing calendars:', error);
      return [];
    }
  }

  //@ts-ignore
  router.get('/check-login', async (req, res) => {
    const userEmail = req.session?.userEmail;
    if (!userEmail) {
      return res.json({ loggedIn: false });
    }
    
    const encryptedCredentials = await redisClient.hget(userEmail, 'password');
    if (!encryptedCredentials) {
      return res.json({ loggedIn: false });
    }
    
    const calendarIdsString = await redisClient.hget(userEmail, 'calendarIds');
    const calendarIds = calendarIdsString ? JSON.parse(calendarIdsString) : [];

    res.json({ 
      loggedIn: true, 
      email: userEmail, 
      hasCalendarId: !!calendarIds,
      calendars: calendarIds
    });
  });

  router.post('/login', async (req, res) => {
    const { email, password } = req.body;
    
    logger.info(`Login attempt for email: ${email}`);
    
    try {
      const service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
      service.Credentials = new WebCredentials(email, password);
      service.Url = new Uri('https://mail.example.com/EWS/Exchange.asmx');
      
      await Folder.Bind(service, WellKnownFolderName.Inbox);
      
      const encryptedCredentials = encrypt(JSON.stringify({ email, password }));

      const calendars = await listCalendars(service, WellKnownFolderName.MsgFolderRoot);

      await redisClient.hset(email, {
        'password': encryptedCredentials,
        'calendarIds': JSON.stringify(calendars),
      });

      req.session!.userEmail = email;
      
      res.json({ loggedIn: true, email, calendars});
    } catch (error) {
      logger.error(`Login failed for ${email}: ${error}`);
      res.status(401).json({ 
        error: 'Invalid credentials',
        details: error instanceof Error ? error.message : 'Unknown error'
      });
    }
  });
  
  router.post('/emails', authMiddleware, async (req: CustomRequest, res) => {
    const { email, password } = req.userCredentials;
    const top = parseInt(req.query.top as string, 10) || 100;

    try {
      const service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
      service.Credentials = new WebCredentials(email, password);
      service.Url = new Uri('https://mail.example.com/EWS/Exchange.asmx');

      const view = new ItemView(top);
      view.PropertySet = new PropertySet(
        BasePropertySet.IdOnly,
        EmailMessageSchema.Subject,
        EmailMessageSchema.DateTimeReceived,
        EmailMessageSchema.From
      );

      const findResults = await service.FindItems(WellKnownFolderName.Inbox, view);

      const emails: EmailItem[] = findResults.Items
        .filter((item): item is EmailMessage => item instanceof EmailMessage)
        .map(item => {
          const dateTimeReceived = new Date(item.DateTimeReceived.ToISOString());
          const formattedDateTime = `${dateTimeReceived.getFullYear()}/${String(dateTimeReceived.getMonth() + 1).padStart(2, '0')}/${String(dateTimeReceived.getDate()).padStart(2, '0')} ${String(dateTimeReceived.getHours()).padStart(2, '0')}:${String(dateTimeReceived.getMinutes()).padStart(2, '0')}`;
          
          // 創建 OWA 鏈接
          // const owaLink = `https://mail.example.com/owa/?ItemID=${item.Id.UniqueId}&exvsurl=1&viewmodel=ReadMessageItem`;
          const owaLink = `https://mail.example.com/owa/?ItemID=${item.Id.UniqueId.replace(/\+/g, '%2B')}&exvsurl=1&viewmodel=ReadMessageItem`;
          
          return {
            subject: item.Subject,
            receivedDate: formattedDateTime,
            from: item.From.Name,
            link: owaLink,
          };
        });

      res.json(emails);
    } catch (error) {
      logger.error('Error fetching emails:'+ error);
      res.status(500).json({ error: 'Failed to fetch emails' });
    }
  });
  
  //@ts-ignore
  router.post('/calendar', authMiddleware, async (req:CustomRequest, res) => {
    const { email, password } = req.userCredentials;
    const { calendarId, timeMin, timeMax } = req.body;
    
    if (!calendarId) {
      return res.status(400).json({ error: 'CalendarId is required' });
    }
  
    try {
      const service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
      service.Credentials = new WebCredentials(email, password);
      service.Url = new Uri('https://mail.example.com/EWS/Exchange.asmx');
  
      const folderId = new FolderId(calendarId);
      const calendarFolder = await Folder.Bind(service, folderId);
      logger.info(`Calendar name: ${calendarFolder.DisplayName}`);
  
      const startDate = timeMin ? DateTime.Parse(timeMin) : DateTime.Now;
      const endDate = timeMax ? DateTime.Parse(timeMax) : startDate.AddDays(1);
      const view = new CalendarView(startDate, endDate);
      view.PropertySet = new PropertySet(
        BasePropertySet.FirstClassProperties,
        AppointmentSchema.Subject,
        AppointmentSchema.Start,
        AppointmentSchema.End,
        AppointmentSchema.Location,
        AppointmentSchema.IsAllDayEvent,
        AppointmentSchema.IsRecurring
      );
  
      const findResults = await service.FindAppointments(folderId, view);
  
      const events: CalendarEvent[] = findResults.Items
        .filter((item): item is Appointment => item instanceof Appointment)
        .map(item => ({
          id: item.Id.UniqueId,
          subject: item.Subject,
          start: item.Start.ToISOString(),
          end: item.End.ToISOString(),
          location: item.Location,
          isAllDay: item.IsAllDayEvent,
          isRecurring: item.IsRecurring,
        }));
  
      res.json({
        calendarName: calendarFolder.DisplayName,
        events: events
      });
    } catch (error) {
      logger.error('Error fetching calendar events: ' + error);
      res.status(500).json({ error: 'Failed to fetch calendar events' });
    }
  });
  
  //@ts-ignore
  router.post('/update-calendar-id', authMiddleware, async (req: CustomRequest, res) => {
    const { calendarId } = req.body;
    const userEmail = req.session?.userEmail;
  
    if (!calendarId) {
      return res.status(400).json({ success: false, error: 'Calendar ID is required' });
    }
  
    if (!userEmail) {
      return res.status(401).json({ success: false, error: 'User not logged in' });
    }
  
    try {
      // 更新Redis中的calendarId
      await redisClient.hset(userEmail, 'calendarId', calendarId);
      
      // 驗證更新是否成功
      const updatedCalendarId = await redisClient.hget(userEmail, 'calendarId');
      
      if (updatedCalendarId !== calendarId) {
        throw new Error('Failed to update calendar ID in Redis');
      }
      
      logger.info(`Calendar ID updated successfully for user: ${userEmail}`);
      res.json({ success: true, message: 'Calendar ID updated successfully' });
    } catch (error) {
      logger.error(`Error updating calendar ID for user ${userEmail}: ${error}`);
      res.status(500).json({ success: false, error: 'Failed to update calendar ID' });
    }
  });

  // @ts-ignore
  // 新增一個 API 端點來添加日曆
  router.post('/add-calendar', authMiddleware, async (req: CustomRequest, res) => {
    const { calendarId, calendarName } = req.body;
    const userEmail = req.session?.userEmail;

    if (!calendarId || !calendarName) {
      return res.status(400).json({ success: false, error: 'Calendar ID and name are required' });
    }

    try {
      const calendarIdsString = await redisClient.hget(userEmail, 'calendarIds');
      let calendarIds = calendarIdsString ? JSON.parse(calendarIdsString) : [];
      calendarIds.push({ id: calendarId, name: calendarName });
      await redisClient.hset(userEmail, 'calendarIds', JSON.stringify(calendarIds));

      res.json({ success: true, message: 'Calendar added successfully' });
    } catch (error) {
      logger.error(`Error adding calendar for user ${userEmail}: ${error}`);
      res.status(500).json({ success: false, error: 'Failed to add calendar' });
    }
  });
  
  // @ts-ignore
  router.post('/delete-calendar', authMiddleware, async (req: CustomRequest, res) => {
    const { calendarId } = req.body;
    const userEmail = req.session?.userEmail;
  
    if (!calendarId) {
      return res.status(400).json({ success: false, error: 'Calendar ID is required' });
    }
  
    try {
      const calendarIdsString = await redisClient.hget(userEmail, 'calendarIds');
      let calendarIds = calendarIdsString ? JSON.parse(calendarIdsString) : [];
      
      // 過濾掉要刪除的日曆
      calendarIds = calendarIds.filter((calendar: { id: any; }) => calendar.id !== calendarId);
      
      // 更新 Redis 中的日曆列表
      await redisClient.hset(userEmail, 'calendarIds', JSON.stringify(calendarIds));
  
      res.json({ success: true, message: 'Calendar deleted successfully' });
    } catch (error) {
      logger.error(`Error deleting calendar for user ${userEmail}: ${error}`);
      res.status(500).json({ success: false, error: 'Failed to delete calendar' });
    }
  });

  router.use(errorHandler());
  return router;
}