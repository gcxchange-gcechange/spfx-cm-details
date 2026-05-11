import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxCmDetailsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  context: WebPartContext;
  hasTeamsContext: boolean;
  userDisplayName: string;
  prefLang: string;
  environment: string;
  devCareerMarketplaceTermSetId: string;
  devJobTypeTermId: string;
  devProgramAreaTermId: string;
  devProgramAreaColumnName: string;
  devAuthClientId: string;
  devDeleteAPIURL: string;
  devCareerMarketplaceHomePage: string;
  devEditOpportunityPage: string;
  uatCareerMarketplaceTermSetId: string;
  uatJobTypeTermId: string;
  uatProgramAreaTermId: string;
  uatProgramAreaColumnName: string;
  uatAuthClientId: string;
  uatDeleteAPIURL: string;
  uatCareerMarketplaceHomePage: string;
  uatEditOpportunityPage: string;
  prodCareerMarketplaceTermSetId: string;
  prodJobTypeTermId: string;
  prodProgramAreaTermId: string;
  prodProgramAreaColumnName: string;
  prodAuthClientId: string;
  prodDeleteAPIURL: string;
  prodCareerMarketplaceHomePage: string;
  prodEditOpportunityPage: string;
}
