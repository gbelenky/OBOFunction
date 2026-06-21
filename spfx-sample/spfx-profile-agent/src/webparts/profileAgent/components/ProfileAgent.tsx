import * as React from 'react';
import styles from './ProfileAgent.module.scss';
import type { IProfileAgentProps } from './IProfileAgentProps';
import { ProxyClient, ChatReply } from '../services/ProxyClient';

interface IChatTurn {
  role: 'user' | 'agent';
  text: string;
}

interface IProfileAgentState {
  input: string;
  chat: IChatTurn[];
  asking: boolean;
  chatError: string | undefined;
  maximized: boolean;
}

export default class ProfileAgent extends React.Component<IProfileAgentProps, IProfileAgentState> {
  private readonly proxy: ProxyClient;

  public constructor(props: IProfileAgentProps) {
    super(props);
    this.proxy = new ProxyClient(props.aadHttpClientFactory);
    this.state = {
      input: '',
      chat: [],
      asking: false,
      chatError: undefined,
      maximized: false
    };
  }

  private toggleMaximize = (): void => {
    this.setState(prev => ({ maximized: !prev.maximized }));
  };

  public componentDidMount(): void {
    this.greet();
  }

  private greet = (): void => {
    this.setState({ asking: true, chatError: undefined });
    // Front-end is agnostic: it does NOT prescribe greeting wording or ask for any profile
    // fields. It only signals "the chat was opened" via greet(); the proxy + agent own the
    // greeting text (one short sentence, by first name, no profile dump).
    this.proxy.greet()
      .then((data: ChatReply) => {
        this.setState(prev => ({
          asking: false,
          chat: [...prev.chat, { role: 'agent', text: data.reply }]
        }));
      })
      .catch((e: Error) => this.setState({ asking: false, chatError: e.message }));
  };

  private ask = (): void => {
    const message = this.state.input.trim();
    if (!message || this.state.asking) {
      return;
    }
    this.setState(prev => ({
      asking: true,
      chatError: undefined,
      input: '',
      chat: [...prev.chat, { role: 'user', text: message }]
    }));
    this.proxy.ask(message)
      .then((data: ChatReply) => {
        this.setState(prev => ({
          asking: false,
          chat: [...prev.chat, { role: 'agent', text: data.reply }]
        }));
      })
      .catch((e: Error) => this.setState({ asking: false, chatError: e.message }));
  };

  private onInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ input: e.target.value });
  };

  private onInputKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter') {
      this.ask();
    }
  };

  public render(): React.ReactElement<IProfileAgentProps> {
    const { hasTeamsContext } = this.props;
    const { input, chat, asking, chatError, maximized } = this.state;
    const sectionClass =
      `${styles.profileAgent} ${hasTeamsContext ? styles.teams : ''} ${maximized ? styles.maximized : ''}`;
    return (
      <section className={sectionClass}>
        <div className={styles.welcome}>
          <h2>Ask the agent</h2>
          <button
            type="button"
            className={styles.maxButton}
            onClick={this.toggleMaximize}
            aria-pressed={maximized}
            title={maximized ? 'Restore' : 'Maximize'}
          >
            {maximized ? '🗗 Restore' : '🗖 Maximize'}
          </button>
        </div>

        <div className={styles.chat}>
          <div className={styles.transcript}>
            {chat.length === 0 ? (
              <p className={styles.hint}>
                {asking ? 'Greeting you…' : 'Try: \u201cHow do I request vacation?\u201d'}
              </p>
            ) : (
              chat.map((t, i) => (
                <div key={i} className={t.role === 'user' ? styles.userTurn : styles.agentTurn}>
                  <strong>{t.role === 'user' ? 'You' : 'Agent'}:</strong>{' '}
                  <span style={{ whiteSpace: 'pre-wrap' }}>{t.text}</span>
                </div>
              ))
            )}
          </div>
          {chatError ? <p className={styles.error}>{chatError}</p> : undefined}
          <div className={styles.inputRow}>
            <input
              type="text"
              className={styles.input}
              placeholder="Ask the Intranet"
              value={input}
              disabled={asking}
              onChange={this.onInputChange}
              onKeyDown={this.onInputKeyDown}
            />
            <button className={styles.button} disabled={asking || !input.trim()} onClick={this.ask}>
              {asking ? 'Asking…' : 'Send'}
            </button>
          </div>
        </div>
      </section>
    );
  }
}
