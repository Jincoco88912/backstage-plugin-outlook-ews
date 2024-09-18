import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  useApi,
  configApiRef,
  identityApiRef,
  storageApiRef,
} from '@backstage/core-plugin-api';
import {
  InfoCard,
  Progress,
  ResponseErrorPanel,
} from '@backstage/core-components';
import {
  makeStyles,
  Typography,
  TextField,
  Box,
  Button,
  Theme,
  Tabs,
  Tab,
  Dialog,
  DialogActions,
  DialogContent,
  DialogContentText,
  DialogTitle,
  IconButton,
} from '@material-ui/core';
import {
  NavigateBefore as PrevIcon,
  NavigateNext as NextIcon,
  Add as AddIcon,
  Close as CloseIcon,
  Share as ShareIcon,
} from '@material-ui/icons';
import { DateTime } from 'luxon';
import { sortBy } from 'lodash';

interface CalendarEvent {
  id: string;
  subject: string;
  start: string;
  end: string;
  location?: string;
  isAllDay: boolean;
  isRecurring: boolean;
}

const useStyles = makeStyles((theme: Theme) => ({
  eventBox: {
    marginBottom: theme.spacing(1),
    padding: theme.spacing(1),
    borderLeft: `4px solid ${theme.palette.primary.main}`,
    '&:hover': {
      backgroundColor: theme.palette.action.hover,
    },
  },
  eventTime: {
    fontSize: '0.8rem',
    color: theme.palette.text.secondary,
  },
  eventSubject: {
    fontWeight: 'bold',
  },
  eventLocation: {
    fontSize: '0.8rem',
    color: theme.palette.text.secondary,
    fontStyle: 'italic',
  },
  loginContainer: {
    padding: theme.spacing(3),
  },
  form: {
    width: '100%',
    marginTop: theme.spacing(1),
  },
  submit: {
    margin: theme.spacing(3, 0, 2),
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: theme.spacing(2),
  },
  dateNavigation: {
    display: 'flex',
    alignItems: 'center',
  },
  addButton: {
    marginLeft: 'auto',
  },
  formContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: theme.spacing(2),
    padding: theme.spacing(3),
    backgroundColor: theme.palette.background.paper,
    borderRadius: theme.shape.borderRadius,
    boxShadow: theme.shadows[1],
  },
  formActions: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: theme.spacing(1),
    marginTop: theme.spacing(2),
  },
  deleteButton: {
    padding: 0,
    marginLeft: theme.spacing(1),
  },
  tabLabel: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    width: '100%',
  },
  tabContent: {
    flexGrow: 1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  helpButton: {
    marginRight: theme.spacing(1),
  },
  calendarIdDisplay: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-start',
    marginTop: theme.spacing(1),
    marginBottom: theme.spacing(1),
  },
  shareButton: {
    marginTop: theme.spacing(1),
    marginBottom: theme.spacing(1),
  },
  calendarIdText: {
    wordBreak: 'break-all',
  },
}));

interface AddCalendarFormProps {
  onClose: () => void;
  onAddCalendar: (newCalendar: { id: string; name: string }) => void;
}

interface Calendar {
  id: string;
  name: string;
}

export const OutlookCalendarComponent = () => {
  const classes = useStyles();
  const [events, setEvents] = useState<CalendarEvent[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error>();
  const [date, setDate] = useState(DateTime.now());
  const [userEmail, setUserEmail] = useState<string>('');
  const [isSignedIn, setIsSignedIn] = useState<boolean>(false);
  const [hasCalendarId, setHasCalendarId] = useState<boolean>(false);
  const [calendars, setCalendars] = useState<Calendar[]>([]);
  const [activeCalendarId, setActiveCalendarId] = useState<string>('');
  const [showAddCalendarForm, setShowAddCalendarForm] = useState(false);
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const [calendarToDelete, setCalendarToDelete] = useState<string | null>(null);
  const [activeCalendarName, setActiveCalendarName] = useState<string>('');
  const [shareDialogOpen, setShareDialogOpen] = useState(false);

  const config = useApi(configApiRef);
  const identityApi = useApi(identityApiRef);
  const storage = useApi(storageApiRef);
  
  const outlookCalendarStore = storage.forBucket('outlookCalendar');

  const handleShareClick = () => {
    setShareDialogOpen(true);
  };

  const handleShareClose = () => {
    setShareDialogOpen(false);
  };

  const handleDeleteClick = (event: React.MouseEvent, calendarId: string) => {
    event.stopPropagation();
    setCalendarToDelete(calendarId);
    setDeleteConfirmOpen(true);
  };

  useEffect(() => {
    const fetchLastViewedCalendar = async () => {
      const lastViewedCalendarId = await outlookCalendarStore.snapshot('lastViewedCalendarId').value;
      if (typeof lastViewedCalendarId === 'string') {
        setActiveCalendarId(lastViewedCalendarId);
      }
    };
    fetchLastViewedCalendar();
  }, [outlookCalendarStore]);

  const handleDeleteConfirm = async () => {
    if (calendarToDelete) {
      try {
        const baseUrl = config.getString('backend.baseUrl');
        const response = await fetch(
          `${baseUrl}/api/outlook-api/delete-calendar`,
          {
            method: 'POST',
            credentials: 'include',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ calendarId: calendarToDelete }),
          },
        );

        if (!response.ok) {
          throw new Error('Failed to delete calendar');
        }

        const data = await response.json();
        if (data.success) {
          setCalendars(prevCalendars =>
            prevCalendars.filter(cal => cal.id !== calendarToDelete),
          );
          if (activeCalendarId === calendarToDelete) {
            setActiveCalendarId(calendars[0]?.id || '');
          }
        } else {
          throw new Error(data.error || 'Failed to delete calendar');
        }
      } catch (error) {
        console.error('Error deleting calendar:', error);
        // 在這裡可以添加用戶友好的錯誤處理,比如顯示一個錯誤提示
      }
    }
    setDeleteConfirmOpen(false);
    setCalendarToDelete(null);
  };

  const handleDeleteCancel = () => {
    setDeleteConfirmOpen(false);
    setCalendarToDelete(null);
  };

  const checkLoginStatus = useCallback(async () => {
    try {
      const baseUrl = config.getString('backend.baseUrl');
      const response = await fetch(`${baseUrl}/api/outlook-api/check-login`, {
        credentials: 'include',
      });
      const data = await response.json();
      if (data.loggedIn) {
        setIsSignedIn(true);
        setUserEmail(data.email);
        setCalendars(data.calendars);
        setHasCalendarId(data.hasCalendarId);
        if (data.calendars.length > 0) {
          const lastViewedCalendarId = await outlookCalendarStore.snapshot('lastViewedCalendarId').value;
          if (typeof lastViewedCalendarId === 'string' && data.calendars.some((cal: Calendar) => cal.id === lastViewedCalendarId)) {
            setActiveCalendarId(lastViewedCalendarId);
          } else {
            setActiveCalendarId(data.calendars[0].id);
          }
        }
      }
    } catch (error) {
      console.error('Failed to check login status:', error);
    }
  }, [config, outlookCalendarStore]);

  useEffect(() => {
    checkLoginStatus();
  }, [checkLoginStatus]);

  useEffect(() => {
    const fetchUserEmail = async () => {
      const { email } = await identityApi.getProfileInfo();
      setUserEmail(email || '');
    };
    fetchUserEmail();
  }, [identityApi]);

  const fetchCalendarEvents = useCallback(
    async (calendarId?: string) => {
      if (!isSignedIn) {
        return;
      }
      const currentCalendarId = calendarId || activeCalendarId;
      if (!currentCalendarId) {
        return;
      }

      setLoading(true);
      try {
        const baseUrl = config.getString('backend.baseUrl');
        const response = await fetch(`${baseUrl}/api/outlook-api/calendar`, {
          method: 'POST',
          credentials: 'include',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            calendarId: activeCalendarId,
            timeMin: date.startOf('day').toISO(),
            timeMax: date.endOf('day').toISO(),
          }),
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Failed to fetch calendar events');
        }
        const data = await response.json();
        if (data.events) {
          setEvents(data.events);
        } else {
          console.warn('No events data in the response');
          setEvents([]);
        }
        setError(undefined);
      } catch (err) {
        console.error('Error fetching calendar events:', err);
        setError(err as Error);
        setEvents([]);
      } finally {
        setLoading(false);
      }
    },
    [config, isSignedIn, hasCalendarId, date, activeCalendarId],
  );

  // @ts-ignore
  useEffect(() => {
    if (isSignedIn && activeCalendarId) {
      fetchCalendarEvents(activeCalendarId);

      const intervalId = setInterval(() => {
        fetchCalendarEvents(activeCalendarId);
      }, 5 * 60 * 1000);
      return () => clearInterval(intervalId);
    }
  }, [isSignedIn, activeCalendarId, fetchCalendarEvents]);

  const handleChangeCalendar = async (
    _event: React.ChangeEvent<{}>,
    newValue: string,
  ) => {
    setActiveCalendarId(newValue);
    const selectedCalendar = calendars.find(cal => cal.id === newValue);
    setActiveCalendarName(selectedCalendar?.name || '');
    setEvents([]);
    fetchCalendarEvents(newValue);
    await outlookCalendarStore.set('lastViewedCalendarId', newValue);
  };

  const changeDay = (offset: number) => {
    setDate(prev => prev.plus({ day: offset }));
  };

  const LoginForm = () => {
    const passwordRef = useRef<HTMLInputElement>(null);
    const emailRef = useRef<HTMLInputElement>(null);

    const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
      const passwordValue = passwordRef.current?.value || '';
      const emailValue = emailRef.current?.value || '';

      event.preventDefault();
      try {
        const baseUrl = config.getString('backend.baseUrl');
        const response = await fetch(`${baseUrl}/api/outlook-api/login`, {
          method: 'POST',
          credentials: 'include',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ email: emailValue, password: passwordValue }),
        });

        if (!response.ok) {
          throw new Error('Login failed');
        }

        const data = await response.json();
        if (data.loggedIn) {
          setIsSignedIn(true);
          setUserEmail(data.email);
          checkLoginStatus();
          window.location.reload();
        }
      } catch (error) {
        console.error('Login error:', error);
        setError(error as Error);
      }
    };

    return (
      <Box className={classes.loginContainer}>
        <Typography component="h1" variant="h5">
          登入同步 Outlook 日曆
        </Typography>
        <form className={classes.form} onSubmit={handleSubmit}>
          <TextField
            variant="outlined"
            margin="normal"
            fullWidth
            id="email"
            label="Email Address"
            name="email"
            inputRef={emailRef}
            defaultValue={userEmail}
            onKeyDown={e => {
              e.stopPropagation();
            }}
          />
          <TextField
            variant="outlined"
            margin="normal"
            fullWidth
            name="password"
            label="Password"
            type="password"
            id="password"
            inputRef={passwordRef}
            onKeyDown={e => {
              e.stopPropagation();
            }}
          />
          <Button
            type="submit"
            fullWidth
            variant="contained"
            color="primary"
            className={classes.submit}
          >
            登入
          </Button>
        </form>
      </Box>
    );
  };

  const AddCalendarForm: React.FC<AddCalendarFormProps> = ({ onClose, onAddCalendar }) => {
    const [calendarId, setCalendarId] = useState('');
    const [calendarName, setCalendarName] = useState('');

    const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
      event.preventDefault();
      try {
        const baseUrl = config.getString('backend.baseUrl');
        const response = await fetch(
          `${baseUrl}/api/outlook-api/add-calendar`,
          {
            method: 'POST',
            credentials: 'include',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ calendarId, calendarName }),
          },
        );

        if (!response.ok) {
          throw new Error('Failed to add calendar');
        }

        const data = await response.json();
        if (data.success) {
          onAddCalendar({ id: calendarId, name: calendarName });
          onClose();
        } else {
          throw new Error(data.error || 'Failed to add calendar');
        }
      } catch (error) {
        console.error('Error adding calendar:', error);
        // Use a more user-friendly error handling method here
      }
    };

    return (
      <form onSubmit={handleSubmit} className={classes.formContainer}>
        <Typography variant="h6">加入行事曆</Typography>
        <TextField
          label="Calendar ID"
          value={calendarId}
          onChange={e => setCalendarId(e.target.value)}
          fullWidth
          variant="outlined"
        />
        <TextField
          label="命名行事曆"
          value={calendarName}
          onChange={e => setCalendarName(e.target.value)}
          fullWidth
          variant="outlined"
        />
        <div className={classes.formActions}>
          <Button onClick={onClose} color="primary">
            取消
          </Button>
          <Button type="submit" variant="contained" color="primary">
            加入
          </Button>
        </div>
      </form>
    );
  };

  return (
    <InfoCard
      deepLink={{
        link: 'https://mail.example.com/owa/#path=/calendar/view/Month',
        title: '前往Outlook',
      }}
    >
      <div className={classes.header}>
        <div className={classes.dateNavigation}>
          <IconButton onClick={() => changeDay(-1)} size="small">
            <PrevIcon />
          </IconButton>
          <Typography variant="h6" style={{ margin: '0 16px' }}>
            {date.toLocaleString({
              weekday: 'short',
              month: 'short',
              day: 'numeric',
            })}
          </Typography>
          <IconButton onClick={() => changeDay(1)} size="small">
            <NextIcon />
          </IconButton>
        </div>
        <Button
          variant="outlined"
          color="primary"
          startIcon={<AddIcon />}
          onClick={() => setShowAddCalendarForm(true)}
          className={classes.addButton}
        >
          新增
        </Button>
        <Button
        variant="outlined"
        color="primary"
        startIcon={<ShareIcon />}
        onClick={handleShareClick}
        className={classes.shareButton}
      >
        分享
      </Button>
      </div>
      <Tabs
        value={activeCalendarId}
        onChange={handleChangeCalendar}
        variant="scrollable"
        scrollButtons="auto"
      >
        {calendars.map(calendar => (
          <Tab
            key={calendar.id}
            value={calendar.id}
            label={
              <span className={classes.tabLabel}>
                <span className={classes.tabContent}>{calendar.name}</span>
                <IconButton
                  size="small"
                  className={classes.deleteButton}
                  onClick={e => {
                    e.stopPropagation();
                    handleDeleteClick(e, calendar.id);
                  }}
                >
                  <CloseIcon fontSize="small" />
                </IconButton>
              </span>
            }
          />
        ))}
      </Tabs>
      <Box>
        {loading && <Progress />}
        {error && <ResponseErrorPanel error={error} />}
        {!isSignedIn ? (
          <LoginForm />
        ) : (
          <>
            {showAddCalendarForm ? (
              <AddCalendarForm
                onClose={() => setShowAddCalendarForm(false)}
                onAddCalendar={(newCalendar: { id: string; name: string }) => {
                  setCalendars(prevCalendars => [...prevCalendars, newCalendar]);
                  setShowAddCalendarForm(false);
                  checkLoginStatus();
                }}
              />
            ) : (
              <Box p={1} pb={0} maxHeight={602} overflow="auto">
                {events.length === 0 && (
                  <Box pt={2} pb={2}>
                    <Typography align="center" variant="h6">
                      無事項
                    </Typography>
                  </Box>
                )}
                {sortBy(events, ['start']).map(event => (
                  <Box
                    key={event.id}
                    className={classes.eventBox}
                    style={
                      event.isAllDay
                        ? { backgroundColor: 'rgba(0, 0, 0, 0.08)' }
                        : {}
                    }
                  >
                    {!event.isAllDay && (
                      <Typography className={classes.eventTime}>
                        {DateTime.fromISO(event.start).toLocaleString(
                          DateTime.TIME_SIMPLE,
                        )}{' '}
                        -
                        {DateTime.fromISO(event.end).toLocaleString(
                          DateTime.TIME_SIMPLE,
                        )}
                      </Typography>
                    )}
                    <Typography
                      className={classes.eventSubject}
                      style={
                        event.isAllDay
                          ? { fontWeight: 'bold', color: '#3f51b5' }
                          : {}
                      }
                    >
                      {event.isAllDay ? '全天：' : ''}
                      {event.subject}
                    </Typography>
                    {event.location && (
                      <Typography className={classes.eventLocation}>
                        {event.location}
                      </Typography>
                    )}
                  </Box>
                ))}
              </Box>
            )}
          </>
        )}
      </Box>
      <Dialog
        open={deleteConfirmOpen}
        onClose={handleDeleteCancel}
        aria-labelledby="alert-dialog-title"
        aria-describedby="alert-dialog-description"
      >
        <DialogTitle id="alert-dialog-title">{'確認刪除行事曆'}</DialogTitle>
        <DialogContent>
          <DialogContentText id="alert-dialog-description">
            您確定要刪除這個行事曆嗎？此操作無法撤銷。
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={handleDeleteCancel} color="primary">
            取消
          </Button>
          <Button
            onClick={() => handleDeleteConfirm()}
            color="primary"
            autoFocus
          >
            確認刪除
          </Button>
        </DialogActions>
      </Dialog>
      <Dialog open={shareDialogOpen} onClose={handleShareClose}>
        <DialogTitle>複製 ID 來分享您當前行事曆 (名稱：{activeCalendarName})</DialogTitle>
        <DialogContent>
          <Typography variant="body1" className={classes.calendarIdText}>       
            {activeCalendarId}
          </Typography>
          <DialogContentText>
            <br></br>
            請先確保您的行事曆是公開的，並且對方已經有檢視權限，然後將此 ID 分享給其他人。
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={handleShareClose} color="primary">
            關閉
          </Button>
        </DialogActions>
      </Dialog>
    </InfoCard>
  );
};

export default OutlookCalendarComponent;