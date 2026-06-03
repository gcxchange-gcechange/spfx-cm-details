/* eslint-disable @typescript-eslint/no-explicit-any */

export interface IEnvConfig {
  careerMarketplaceTermSetId: string;
  jobTypeTermId: string;
  programAreaTermId: string;
  programAreaColumnName: string;
  jobTypeColumnName: string;
  authClientId: string;
  deleteAPIURL: string;
  careerMarketplaceHomePage: string;
  editOpportunityPage: string;
}

export const getEnvConfig = (environment: string, props: any): IEnvConfig => {
  const configs: Record<string, IEnvConfig> = {
     dev: {
      careerMarketplaceTermSetId: props.devCareerMarketplaceTermSetId,
      jobTypeTermId: props.devJobTypeTermId,
      programAreaTermId: props.devProgramAreaTermId,
      programAreaColumnName: props.devProgramAreaColumnName,
      jobTypeColumnName: props.devJobTypeColumnName,
      authClientId: props.devAuthClientId,
      deleteAPIURL: props.devDeleteAPIURL,
      careerMarketplaceHomePage: props.devCareerMarketplaceHomePage,
      editOpportunityPage: props.devEditOpportunityPage
    },
    uat: {
      careerMarketplaceTermSetId: props.uatCareerMarketplaceTermSetId,
      jobTypeTermId: props.uatJobTypeTermId,
      programAreaTermId: props.uatProgramAreaTermId,
      programAreaColumnName: props.uatProgramAreaColumnName,
      jobTypeColumnName: props.uatJobTypeColumnName,
      authClientId: props.uatAuthClientId,
      deleteAPIURL: props.uatDeleteAPIURL,
      careerMarketplaceHomePage: props.uatCareerMarketplaceHomePage,
      editOpportunityPage: props.uatEditOpportunityPage,
    },
   prod:{
      careerMarketplaceTermSetId: props.prodCareerMarketplaceTermSetId,
      jobTypeTermId: props.prodJobTypeTermId,
      programAreaTermId: props.prodProgramAreaTermId,
      programAreaColumnName: props.prodProgramAreaColumnName,
      jobTypeColumnName: props.prodJobTypeColumnName,
      authClientId: props.prodAuthClientId,
      deleteAPIURL: props.prodDeleteAPIURL,
      careerMarketplaceHomePage: props.prodCareerMarketplaceHomePage,
      editOpportunityPage: props.prodEditOpportunityPage,
   }
  };

  const config = configs[environment || 'dev']; // Default to 'dev' if environment is not specified

  if (!config) {
    throw new Error(`Unknown environment: ${environment}, defaulting to Dev`);
  }

  return config;
};