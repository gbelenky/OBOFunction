import { AadHttpClientFactory } from '@microsoft/sp-http';

export interface IProfileAgentProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  aadHttpClientFactory: AadHttpClientFactory;
}
