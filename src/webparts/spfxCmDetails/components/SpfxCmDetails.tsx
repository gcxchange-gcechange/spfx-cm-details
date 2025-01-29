/* eslint-disable dot-notation */
import * as React from 'react';
import styles from './SpfxCmDetails.module.scss';
import type { ISpfxCmDetailsProps } from './ISpfxCmDetailsProps';
import { getSP } from '../../../pnpConfig';
import { SPFI } from '@pnp/sp';
import { SPFx, graphfi } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/taxonomy";
import { TermStore } from '@microsoft/microsoft-graph-types';
//import { ITermSet } from "@pnp/graph/taxonomy";
import { SelectLanguage } from "./SelectLanguage";
import { PrimaryButton } from '@fluentui/react';

export interface ISpfxCmDetailsState {
    TitleFr: string;
    TitleEn: string;
    DescEn: string;
    DescFr: string;
    JobType: any;
    JobTypeFr: any;
    AppDeadline: string;
    program: any;
    Department: any;
    Nmb_opt: string;
    Duration: any;
    DurationQuantity: string;
    Work_Arr: any;
    Location: any;
    sec_lvl: any;
    Language: any;
    LanguageComprehension: string;
    NoOpt: boolean;
    ContactEmail: string;
    OptId: number;
}
export default class SpfxCmDetails extends React.Component<ISpfxCmDetailsProps, ISpfxCmDetailsState> {

    public strings = SelectLanguage(this.props.prefLang);

    constructor(props: ISpfxCmDetailsProps, state: ISpfxCmDetailsState) {
        super(props);
        this.state = {
            TitleFr: "",
            TitleEn: "",
            DescEn: "",
            DescFr: "",
            JobType: [],
            JobTypeFr: [],
            AppDeadline: "",
            program: [],
            Department: [],
            Nmb_opt: "",
            Duration: [],
            DurationQuantity: "",
            Work_Arr: [],
            Location: [],
            sec_lvl: [],
            Language: [],
            LanguageComprehension: "",
            NoOpt: true,
            ContactEmail: "",
            OptId: 0
        }
    }

    public async componentDidMount(): Promise<void> {
        await this._geturlID();
    }

    public _geturlID = async (): Promise<void> => {
        const params = new URLSearchParams(window.location.search);
        const val = params.get('JobOpportunityId') as unknown as number; // convert to number
        if (val !== null && val ) { // check if value exist and not empty
            this.setState({
                NoOpt: false,
                OptId: val
            })
            await this._getdetailsopt(val);
        } else {
            this.setState({
                NoOpt: true
            })
        }
    }

    public _getdetailsopt = async (valueid: number): Promise<void> => {
        const job_type_termset_ID = "45f37f08-3ff4-4d84-bf21-4a77ddffcf3e";
        const program_area_termset_ID = "bd807536-d8e7-456b-aab0-fae3eecedd8a";

        const _sp: SPFI = getSP(this.props.context);

        try {

            const item = await _sp.web.lists.getByTitle("JobOpportunity").items.getById(valueid).select("Department", "Department/NameEn", "Department/NameFr", "ClassificationCode", "ClassificationCode/NameEn", "ClassificationCode/NameFr", "NumberOfOpportunities", "JobTitleFr", "JobTitleEn", "JobDescriptionEn", "JobDescriptionFr", "ApplicationDeadlineDate", "ContactEmail", "ProgramArea", "JobType", "Duration", "Duration/NameEn", "Duration/NameFr", "DurationQuantity", "WorkArrangement", "WorkArrangement/NameEn", "WorkArrangement/NameFr", "City", "City/NameEn", "City/NameFr", "SecurityClearance", "SecurityClearance/NameEn", "SecurityClearance/NameFr", "LanguageRequirement", "LanguageRequirement/NameEn", "LanguageRequirement/NameFr", "LanguageComprehension").expand("Department", "ClassificationCode", "Duration", "WorkArrangement", "City", "SecurityClearance", "LanguageRequirement")();
            console.log(item);

            this.setState({
                TitleFr: item.JobTitleFr,
                TitleEn: item.JobTitleEn,
                DescEn: item.JobDescriptionEn,
                DescFr: item.JobDescriptionFr,
                JobType: await this._get_terms(job_type_termset_ID,item.JobType[0].TermGuid),
                program: await this._get_terms(program_area_termset_ID, item.ProgramArea.TermGuid),
                Department: item.Department,
                AppDeadline: item.ApplicationDeadlineDate.split('T')[0], // convert into format YYYY/MM/DD
                Nmb_opt: item.NumberOfOpportunities,
                Duration: item.Duration,
                DurationQuantity: item.DurationQuantity,
                Work_Arr:item.WorkArrangement,
                Location: item.City,
                sec_lvl: item.SecurityClearance,
                Language: item.LanguageRequirement,
                ContactEmail: item.ContactEmail,
                LanguageComprehension: item.LanguageComprehension
            })
        } catch(e) {
            
            console.error(e);
            this.setState({
                NoOpt: true
            })
        }
    }

    public _get_terms = async (termsetid: string, termsid: string): Promise<void> => {

        const graph = graphfi().using(SPFx(this.props.context));

        let lang_id = 0;
        if (this.props.prefLang === "fr-fr") {
            lang_id = 1;
        } else {
            lang_id = 0;
        }

        const info: TermStore.Term = await graph.termStore.groups.getById("656c725c-def6-46cd-86df-b51f1b22383e").sets.getById(termsetid).getTermById(termsid)();
        return JSON.parse(JSON.stringify(info.labels))[lang_id].name;
    }


  public render(): React.ReactElement<ISpfxCmDetailsProps> {
    const {
      hasTeamsContext,
    } = this.props;

    return (
        <section className={`${styles.spfxCmDetails} ${hasTeamsContext ? styles.teams : ''}`}>
            {this.state.NoOpt ? (

                <div className={styles.welcome}>
                    <h2>Sorry! This opportunity do not exist+{this.state.NoOpt}</h2>
                    <p>Please, try another one or reach out to our support team!</p>
                </div>

            ) : ( 

           <>
                <div className={styles.welcome}>
                             <h2>{this.props.prefLang === "fr-fr" ? (
                                this.state.TitleFr
                            ) : (
                                this.state.TitleEn
                            )}</h2>
                </div>
                        <div>
                            <p className={styles.desc_bold}>
                                {this.props.prefLang === "fr-fr" ? (
                                    this.state.DescFr
                                ) : (
                                    this.state.DescEn
                                )}
                            </p>
                            <div className={styles.deadline_type_section}>
                        <span className={styles.jobtype_space}>Job type ({this.state.JobType})</span>
                        <span>Application deadline: {this.state.AppDeadline}</span>
                    </div>
                <div>
                    <h3>Opportunity Details:</h3>
                    <p>
                        <h4>Program area</h4>
                        {this.state.program}
                    </p>
                    <p>
                    <h4>Department</h4>
                        {this.props.prefLang === "fr-fr" ? (
                            this.state.Department.NameFr
                        ) : (
                            this.state.Department.NameEn
                        )}
                    </p>
                    <p>
                        <h4>Number of opportunities</h4>
                        {this.state.Nmb_opt}
                    </p>
                    <p>
                        <h4>Duration</h4>
                        {this.state.DurationQuantity + " "}
                        {this.props.prefLang === "fr-fr" ? (
                            this.state.Duration.NameFr
                        ) : (
                            this.state.Duration.NameEn
                        )}
                       
                    </p>
                    <p>
                        <h4>Application deadline</h4>
                        {this.state.AppDeadline}
                    </p>
                    <p>
                        <h4>Work arrangement</h4>
                        {this.props.prefLang === "fr-fr" ? (
                            this.state.Work_Arr.NameFr
                        ) : (
                            this.state.Work_Arr.NameEn
                        )}
                    </p>
                    <p>
                        <h4>Location</h4>
                        {this.props.prefLang === "fr-fr" ? (
                            this.state.Location.NameFr
                        ) : (
                            this.state.Location.NameEn
                        )}
                    </p>
                    <p>
                        <h4>Security level</h4>
                        {this.props.prefLang === "fr-fr" ? (
                            this.state.sec_lvl.NameFr
                        ) : (
                            this.state.sec_lvl.NameEn
                        )}
                        
                    </p>
                    <p>
                        <h4>Language requirements</h4>
                        {this.props.prefLang === "fr-fr" ? (
                            this.state.Language.NameFr
                        ) : (
                            this.state.Language.NameEn
                        )}
                        { this.state.LanguageComprehension }
                    </p>
                            </div>
                            <PrimaryButton text="Apply" href={`mailto: ${this.state.ContactEmail}?subject=The%20subject%20of%20the%20mail&body=The%20body%20of%20the%20email&?JobOpportunityId=${this.state.OptId}`} />
               </div>
             </>    
            )
        }
      </section>
    );
  }
}
