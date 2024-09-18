import React, { useState, useEffect, useRef, useCallback } from 'react';
import { useApi, configApiRef, identityApiRef } from '@backstage/core-plugin-api';
import {
  Table,
  Progress,
  ResponseErrorPanel,
  TableColumn,
} from '@backstage/core-components';
import { 
  Button, 
  Chip, 
  makeStyles, 
  Typography, 
  TextField,
  Paper,
  IconButton,
} from '@material-ui/core';
import { Refresh as RefreshIcon, OpenInNew as OpenInNewIcon } from '@material-ui/icons';

interface EmailItem {
  subject: string;
  receivedDate: string;
  from: string;
  link: string;
}

const useStyles = makeStyles(theme => ({
  container: {
    width: '100%',
  },
  loginContainer: {
    padding: theme.spacing(3),
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    marginBottom: theme.spacing(3),
  },
  form: {
    width: '100%',
    marginTop: theme.spacing(1),
  },
  submit: {
    margin: theme.spacing(3, 0, 2),
  },
  titleContainer: {
    display: 'flex',
    alignItems: 'center',
  },
}));

const columns: TableColumn[] = [
  {
    title: '標題',
    field: 'subject',
    highlight: true,
    render: (row: Partial<EmailItem>) => (
      // @ts-ignore
      <Button
        color="primary" 
        href={row.link} 
        target="_blank" 
        rel="noopener noreferrer"
      >
        {row.subject}
      </Button>
    ),
  },
  {
    title: '寄件者',
    field: 'from',
    render: (row: Partial<EmailItem>) => (
      <Chip label={row.from} variant="outlined" size="small" />
    ),
  },
  { title: '寄信日期', field: 'receivedDate' },
];

export const OutlookComponent = () => {
  const classes = useStyles();
  const [emails, setEmails] = useState<EmailItem[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<Error>();
  const [userEmail, setUserEmail] = useState<string>('');
  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(false);
  const config = useApi(configApiRef);
  const identityApi = useApi(identityApiRef);

  useEffect(() => {
    const fetchUserEmail = async () => {
      const { email } = await identityApi.getProfileInfo();
      setUserEmail(email || '');
    };
    fetchUserEmail();
  }, [identityApi]);

  const checkLoginStatus = useCallback(async () => {
    try {
      const baseUrl = config.getString('backend.baseUrl');
      const response = await fetch(`${baseUrl}/api/outlook-api/check-login`, {
        credentials: 'include',
      });
      const data = await response.json();
      if (data.loggedIn) {
        setIsLoggedIn(true);
        setUserEmail(data.email);
      }
    } catch (error) {
      console.error('Failed to check login status:', error);
    }
  }, [config]);
  
  useEffect(() => {
    checkLoginStatus();
  }, [checkLoginStatus]);
  
  const fetchEmails = useCallback(async () => {
    if (!isLoggedIn) {
      setError(new Error('Please log in first'));
      return;
    }
  
    setLoading(true);
    try {
      const baseUrl = config.getString('backend.baseUrl');
      const response = await fetch(`${baseUrl}/api/outlook-api/emails`, {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/json',
        },
      });
  
      if (!response.ok) {
        throw new Error('Failed to fetch emails');
      }
      const fetchedEmails = await response.json();
      setEmails(fetchedEmails);
      setError(undefined);
    } catch (err) {
      setError(err as Error);
    } finally {
      setLoading(false);
    }
  }, [config, isLoggedIn]);

  const LoginForm = () => {
    const passwordRef = useRef<HTMLInputElement>(null);
    const emailRef = useRef<HTMLInputElement>(null);

    const handleSubmit = async (event: React.FormEvent<HTMLFormElement>) => {
      event.preventDefault();
      const passwordValue = passwordRef.current?.value || '';
      const emailValue = emailRef.current?.value || '';

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
          setIsLoggedIn(true);
          setUserEmail(data.email);
          window.location.reload();
        }
      } catch (error) {
        console.error('Login error:', error);
        setError(error as Error);
      }
    };
    
    return (
      <Paper className={classes.loginContainer}>
        <Typography component="h1" variant="h5">
          登入同步 Outlook 信件
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
            onKeyDown={(e) => {
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
            onKeyDown={(e) => {
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
            Sign In
          </Button>
        </form>
      </Paper>
    );
  };

  const RefreshButton = () => (
    <Button
      color="primary"
      variant="outlined"
      onClick={fetchEmails}
      startIcon={<RefreshIcon />}
    >
      Refresh
    </Button>
  );

  const TableTitle = () => (
    <div className={classes.titleContainer}>
      <Typography variant="h6">Outlook 信件</Typography>
      <IconButton
        // @ts-ignore
        className={classes.titleIcon}
        href="https://mail.example.com"
        target="_blank"
        rel="noopener noreferrer"
      >
        <OpenInNewIcon fontSize="small" />
      </IconButton>
    </div>
  );

  // @ts-ignore
  useEffect(() => {
    if (isLoggedIn) {
      fetchEmails();

      // Set up an interval to fetch data every 5 minutes
      const intervalId = setInterval(() => {
        fetchEmails();
      }, 5 * 60 * 1000); // 5 minutes in milliseconds
      // Clean up the interval when the component unmounts
      return () => clearInterval(intervalId);

    }
  }, [isLoggedIn, fetchEmails]);

  return (
    <div className={classes.container}>
      {loading && <Progress />}
      {error && <ResponseErrorPanel error={error} />}
      {!isLoggedIn ? (
        <LoginForm />
      ) : (
        <Table
          options={{
            search: true,
            paging: true,
            pageSize: 5,
            padding: 'dense',
          }}
          data={emails}
          columns={columns}
          title={<TableTitle />}
          actions={[
            {
              icon: () => <RefreshButton />,
              tooltip: 'Refresh',
              isFreeAction: true,
              onClick: fetchEmails,
            },
          ]}
        />
      )}
    </div>
  );
};

export default OutlookComponent;