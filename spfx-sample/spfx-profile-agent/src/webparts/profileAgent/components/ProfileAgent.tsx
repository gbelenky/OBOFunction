import * as React from 'react';
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  makeStyles,
  mergeClasses,
  tokens,
  Text,
  Title3,
  Textarea,
  Button,
  Spinner,
  type TextareaProps
} from '@fluentui/react-components';
import {
  Send24Regular,
  FullScreenMaximize24Regular,
  FullScreenMinimize24Regular
} from '@fluentui/react-icons';
import type { IProfileAgentProps } from './IProfileAgentProps';
import { ProxyClient, ChatReply } from '../services/ProxyClient';

interface IChatTurn {
  role: 'user' | 'agent';
  text: string;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    maxWidth: '900px',
    marginLeft: 'auto',
    marginRight: 'auto',
    padding: tokens.spacingVerticalM,
    boxSizing: 'border-box',
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1
  },
  maximized: {
    position: 'fixed',
    inset: '0',
    zIndex: 1000,
    maxWidth: 'none',
    height: '100vh',
    boxShadow: tokens.shadow64
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM
  },
  transcript: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    flexGrow: 1,
    minHeight: '120px',
    overflowY: 'auto',
    paddingRight: tokens.spacingHorizontalXS
  },
  bubble: {
    maxWidth: '80%',
    paddingTop: tokens.spacingVerticalS,
    paddingBottom: tokens.spacingVerticalS,
    paddingLeft: tokens.spacingHorizontalM,
    paddingRight: tokens.spacingHorizontalM,
    borderRadius: tokens.borderRadiusLarge,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word',
    lineHeight: tokens.lineHeightBase300
  },
  userBubble: {
    alignSelf: 'flex-end',
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand
  },
  agentBubble: {
    alignSelf: 'flex-start',
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1
  },
  hint: {
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic'
  },
  error: {
    color: tokens.colorPaletteRedForeground1
  },
  inputRow: {
    display: 'flex',
    alignItems: 'flex-end',
    gap: tokens.spacingHorizontalS
  },
  input: {
    flexGrow: 1
  }
});

/**
 * Chat surface built with the public, supported Fluent UI v9 library
 * (@fluentui/react-components + @fluentui/react-icons).
 *
 * The front-end stays agnostic of the agent: it only talks to the BFF proxy via ProxyClient.
 * It signals "chat opened" through greet() and lets the proxy + hosted agent own greeting/answers.
 */
const ProfileAgent: React.FC<IProfileAgentProps> = (props) => {
  const styles = useStyles();
  const proxyRef = React.useRef<ProxyClient | undefined>(undefined);
  if (!proxyRef.current) {
    proxyRef.current = new ProxyClient(props.aadHttpClientFactory);
  }
  const proxy = proxyRef.current;

  const [input, setInput] = React.useState<string>('');
  const [chat, setChat] = React.useState<IChatTurn[]>([]);
  const [asking, setAsking] = React.useState<boolean>(false);
  const [chatError, setChatError] = React.useState<string | undefined>(undefined);
  const [maximized, setMaximized] = React.useState<boolean>(false);

  const transcriptRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (transcriptRef.current) {
      transcriptRef.current.scrollTop = transcriptRef.current.scrollHeight;
    }
  }, [chat, asking]);

  const runTurn = React.useCallback((call: Promise<ChatReply>): void => {
    setAsking(true);
    setChatError(undefined);
    call
      .then((data) => {
        setAsking(false);
        setChat((prev) => [...prev, { role: 'agent', text: data.reply }]);
      })
      .catch((e: Error) => {
        setAsking(false);
        setChatError(e.message);
      });
  }, []);

  // Greet once when the web part mounts.
  React.useEffect(() => {
    runTurn(proxy.greet());
  }, []);

  const ask = React.useCallback((): void => {
    const message = input.trim();
    if (!message || asking) {
      return;
    }
    setInput('');
    setChat((prev) => [...prev, { role: 'user', text: message }]);
    runTurn(proxy.ask(message));
  }, [input, asking, proxy, runTurn]);

  const onInputChange: TextareaProps['onChange'] = (_e, data) => {
    setInput(data.value);
  };

  const onInputKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>): void => {
    // Enter sends, Shift+Enter inserts a newline.
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      ask();
    }
  };

  return (
    <FluentProvider theme={props.isDarkTheme ? webDarkTheme : webLightTheme}>
      <section className={mergeClasses(styles.root, maximized && styles.maximized)}>
        <div className={styles.header}>
          <Title3>Ask the agent</Title3>
          <Button
            appearance="subtle"
            icon={maximized ? <FullScreenMinimize24Regular /> : <FullScreenMaximize24Regular />}
            onClick={() => setMaximized((v) => !v)}
            aria-pressed={maximized}
            title={maximized ? 'Restore' : 'Maximize'}
          >
            {maximized ? 'Restore' : 'Maximize'}
          </Button>
        </div>

        <div className={styles.transcript} ref={transcriptRef}>
          {chat.length === 0 && !asking ? (
            <Text className={styles.hint}>Try: &ldquo;How do I request vacation?&rdquo;</Text>
          ) : (
            chat.map((t, i) => (
              <div
                key={i}
                className={mergeClasses(
                  styles.bubble,
                  t.role === 'user' ? styles.userBubble : styles.agentBubble
                )}
              >
                {t.text}
              </div>
            ))
          )}
          {asking ? <Spinner size="tiny" label="Thinking…" labelPosition="after" /> : undefined}
        </div>

        {chatError ? <Text className={styles.error}>{chatError}</Text> : undefined}

        <div className={styles.inputRow}>
          <Textarea
            className={styles.input}
            placeholder="Ask the Intranet"
            value={input}
            disabled={asking}
            resize="vertical"
            onChange={onInputChange}
            onKeyDown={onInputKeyDown}
          />
          <Button
            appearance="primary"
            icon={<Send24Regular />}
            disabled={asking || !input.trim()}
            onClick={ask}
          >
            Send
          </Button>
        </div>
      </section>
    </FluentProvider>
  );
};

export default ProfileAgent;
