import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxCmDetailsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  context: WebPartContext;
  hasTeamsContext: boolean;
  userDisplayName: string;
  prefLang: string;
}
