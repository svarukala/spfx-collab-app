import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpFxCollabAppProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spoContext: WebPartContext;
}
