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
    Work_Arr: any;
    Location: any;
    sec_lvl: any;
    Language: any;
    NoOpt: boolean;
    ContactEmail: string;
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
            Work_Arr: [],
            Location: [],
            sec_lvl: [],
            Language: [],
            NoOpt: true,
            ContactEmail: ""
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
                NoOpt: false
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
        const duration_termset_ID = "ad38f4b6-8aec-4e41-b30b-04c56a2aeeb3";
        const program_area_termset_ID = "bd807536-d8e7-456b-aab0-fae3eecedd8a";
        //const class_code_termset_ID = "cc00fcc8-4731-4165-a22d-006ddb7b32ce";
        //const work_schedule_termset_ID = "5a826701-8d58-4c1f-9558-22ea6a98f55f";
        const security_clar_termset_ID = "31a56cc4-eed9-4229-a6b4-d2fdde94f9e5";
        const language_termset_ID = "b1048b91-a228-4425-b728-da90be459f27";
        const work_arr_termset_ID = "74af42a2-246a-41aa-b4bb-8403134f0728";
        const location_termset_ID = "c6d27982-3d09-43d7-828d-daf6e06be362";
        const department_termset_ID = "e86e736d-77a4-447c-8aee-b714be2f64cf";

        const _sp: SPFI = getSP(this.props.context);
        try {
            const items = await _sp.web.lists.getByTitle("JobOpportunityTest").items.getById(valueid)();
        
            console.log(items);

            this.setState({
                TitleFr: items.JobTitleFrTest,
                TitleEn: items.JobTitleEnTest,
                DescEn: items.JobDescriptionEnTest,
                DescFr: items.JobDescriptionFrTest,
                JobType: await this._get_terms(job_type_termset_ID,items.JobTypeTest[0].TermGuid),
                program: await this._get_terms(program_area_termset_ID, items.ProgramAreaTest.TermGuid),
                Department: await this._get_terms(department_termset_ID, items.DepartmentTest.TermGuid),
                AppDeadline: items.ApplicationDeadlineDateTest,
                Nmb_opt: items.NumberOfOpportunitiesTest,
                Duration: await this._get_terms(duration_termset_ID, items.DurationTesst.TermGuid),
                Work_Arr: await this._get_terms(work_arr_termset_ID, items.WorkArrangementTest.TermGuid),
                Location: await this._get_terms(location_termset_ID, items.LocationTest.TermGuid),
                sec_lvl: await this._get_terms(security_clar_termset_ID, items.SecurityClearanceTest.TermGuid),
                Language: await this._get_terms(language_termset_ID, items.LanguageRequirementTest.TermGuid),
                ContactEmail: items.ContactEmailTest
            });

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
                    <h2>{this.state.TitleEn}</h2>
                </div>
                        <div>
                            <p className={styles.desc_bold}>
                        {this.state.DescEn}
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
                        {this.state.Department}
                    </p>
                    <p>
                        <h4>Number of opportunities</h4>
                        {this.state.Nmb_opt}
                    </p>
                    <p>
                        <h4>Duration</h4>
                        {this.state.Duration}
                    </p>
                    <p>
                        <h4>Application deadline</h4>
                        {this.state.AppDeadline}
                    </p>
                    <p>
                        <h4>Work arrangement</h4>
                        {this.state.Work_Arr}
                    </p>
                    <p>
                        <h4>Location</h4>
                        {this.state.Location}
                    </p>
                    <p>
                        <h4>Security level</h4>
                        {this.state.sec_lvl}
                    </p>
                    <p>
                        <h4>Language requirements</h4>
                        {this.state.Language}
                    </p>
                </div>
                <PrimaryButton text="Apply" href={`mailto: ${this.state.ContactEmail}?subject=The%20subject%20of%20the%20mail&body=The%20body%20of%20the%20email`}  />
               </div>
             </>    
            )
        }
      </section>
    );
  }
}
