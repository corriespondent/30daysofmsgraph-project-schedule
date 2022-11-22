import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISchedulerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  context: WebPartContext;
  userDisplayName: string;
}
