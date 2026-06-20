import * as React from 'react';
import styles from './ProfileAgent.module.scss';
import type { IProfileAgentProps } from './IProfileAgentProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ProxyClient, UserProfile, ChatReply } from '../services/ProxyClient';

interface IChatTurn {
  role: 'user' | 'agent';
  text: string;
}

interface IProfileAgentState {
  profile: UserProfile | undefined;
  profileError: string | undefined;
  profileLoading: boolean;
  input: string;
  chat: IChatTurn[];
  asking: boolean;
  chatError: string | undefined;
  consentUrl: string | undefined;
}

export default class ProfileAgent extends React.Component<IProfileAgentProps, IProfileAgentState> {
  private readonly proxy: ProxyClient;

  public constructor(props: IProfileAgentProps) {
    super(props);
    this.proxy = new ProxyClient(props.aadHttpClientFactory);
    this.state = {
      profile: undefined,
      profileError: undefined,
      profileLoading: false,
      input: '',
      chat: [],
      asking: false,
      chatError: undefined,
      consentUrl: undefined
    };
  }

  public componentDidMount(): void {
    this.loadProfile();
  }

  private loadProfile = (): void => {
    this.setState({ profileLoading: true, profileError: undefined });
    this.proxy.loadProfile()
      .then(profile => this.setState({ profile, profileLoading: false }))
      .catch((e: Error) => this.setState({ profileError: e.message, profileLoading: false }));
  };

  private ask = (): void => {
    const message = this.state.input.trim();
    if (!message || this.state.asking) {
      return;
    }
    this.setState(prev => ({
      asking: true,
      chatError: undefined,
      consentUrl: undefined,
      input: '',
      chat: [...prev.chat, { role: 'user', text: message }]
    }));
    this.proxy.ask(message)
      .then((data: ChatReply) => {
        if (data.status === 'consent_required' && data.consentUrl) {
          this.setState({ asking: false, consentUrl: data.consentUrl });
          return;
        }
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

  private renderProfile(): React.ReactNode {
    const { profile, profileError, profileLoading } = this.state;
    if (profileLoading) {
      return <p>Loading your profile…</p>;
    }
    if (profileError) {
      return <p className={styles.error}>Could not load profile: {escape(profileError)}</p>;
    }
    if (!profile) {
      return undefined;
    }
    const extended = profile.sharePointProfile?.extendedProperties ?? profile.extendedProperties ?? {};
    const extendedKeys = Object.keys(extended);
    return (
      <div className={styles.profileCard}>
        <h3>{escape(profile.displayName)}</h3>
        <dl className={styles.profileGrid}>
          {profile.jobTitle ? (<><dt>Title</dt><dd>{escape(profile.jobTitle)}</dd></>) : undefined}
          {profile.department ? (<><dt>Department</dt><dd>{escape(profile.department)}</dd></>) : undefined}
          {profile.mail ? (<><dt>Mail</dt><dd>{escape(profile.mail)}</dd></>) : undefined}
          {profile.officeLocation ? (<><dt>Office</dt><dd>{escape(profile.officeLocation)}</dd></>) : undefined}
          {profile.resolvedVia ? (<><dt>Resolved via</dt><dd>{escape(profile.resolvedVia)}</dd></>) : undefined}
        </dl>
        {extendedKeys.length > 0 ? (
          <>
            <h4>SharePoint profile properties</h4>
            <dl className={styles.profileGrid}>
              {extendedKeys.map(k => (
                <React.Fragment key={k}>
                  <dt>{escape(k)}</dt>
                  <dd>{escape(String(extended[k]))}</dd>
                </React.Fragment>
              ))}
            </dl>
          </>
        ) : undefined}
      </div>
    );
  }

  public render(): React.ReactElement<IProfileAgentProps> {
    const { hasTeamsContext } = this.props;
    const { input, chat, asking, chatError, consentUrl } = this.state;
    return (
      <section className={`${styles.profileAgent} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Your profile &amp; agent</h2>
          <div>Signed in as <strong>{escape(this.props.userDisplayName)}</strong></div>
        </div>

        {this.renderProfile()}

        <div className={styles.chat}>
          <h3>Ask the agent</h3>
          <div className={styles.transcript}>
            {chat.length === 0 ? (
              <p className={styles.hint}>Try: &quot;Who am I and what are my skills?&quot;</p>
            ) : (
              chat.map((t, i) => (
                <div key={i} className={t.role === 'user' ? styles.userTurn : styles.agentTurn}>
                  <strong>{t.role === 'user' ? 'You' : 'Agent'}:</strong> {escape(t.text)}
                </div>
              ))
            )}
          </div>
          {consentUrl ? (
            <p className={styles.error}>
              The agent needs your consent.{' '}
              <a href={consentUrl} target="_blank" rel="noreferrer">Grant consent</a>, then ask again.
            </p>
          ) : undefined}
          {chatError ? <p className={styles.error}>{escape(chatError)}</p> : undefined}
          <div className={styles.inputRow}>
            <input
              type="text"
              className={styles.input}
              placeholder="Ask about your profile…"
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
