/* eslint-disable @typescript-eslint/no-explicit-any */
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
import { PrimaryButton, DefaultButton, IconButton, Icon, Modal } from '@fluentui/react';
import * as strings from 'SpfxCmDetailsWebPartStrings';
import parse from 'html-react-parser';
import { AadHttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';

// @ts-expect-error need this for some reason, * won't work.
import createDOMPurify from 'dompurify';
const DOMPurify = createDOMPurify(window);

export interface ISpfxCmDetailsState {
    TitleFr: string;
    TitleEn: string;
    DescEn: string;
    DescFr: string;
    JobType: any;
    JobTypeFr: any;
    AppDeadline: string;
    program: any;
    classification: any;
    Department: any;
    Nmb_opt: string;
    Duration: any;
    DurationQuantity: string;
    Work_Arr: any;
    LocationEn: string;
    LocationFr: string;
    sec_lvl: any;
    Language: any;
    LanguageComprehension: string;
    NoOpt: boolean;
    ContactEmail: string;
    ContactName: string;
    OptId: number;
    Expired: boolean;
    pageLoading: boolean;
    deleteLoading: boolean;
    deleted: boolean;
    modalOpen: boolean;
    skills: any;
    WorkSchedule: any;
}
export default class SpfxCmDetails extends React.Component<ISpfxCmDetailsProps, ISpfxCmDetailsState> {

    public strings = SelectLanguage(this.props.prefLang);

    /*
        REPLACE THESE FOR YOUR BUILD
    */
    private env = {
        careerMarketplaceTermSetId: '656c725c-def6-46cd-86df-b51f1b22383e',
        jobTypeTermSetId: '45f37f08-3ff4-4d84-bf21-4a77ddffcf3e',
        programAreaTermSetId: 'bd807536-d8e7-456b-aab0-fae3eecedd8a',
        authClientId: 'c121f403-ff41-4db3-8426-f3b9c5016cd4',
        deleteApiUrl: 'https://appsvc-function-dev-cm-listmgmt-dotnet001.azurewebsites.net/api/DeleteJobOpportunity?',
        careerMarketplaceHomePage: 'https://devgcx.sharepoint.com/sites/CM-test',
        editOpportunityPage: 'https://devgcx.sharepoint.com/sites/CM-test/SitePages/editOpportunity.aspx?JobOpportunityId='
    }

    private envValid(): boolean {
        return Object.keys(this.env).some(key => {
            const value = this.env[key as keyof typeof this.env];
            return value === '' || value === null || value === undefined;
        });
    }

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
            classification: "",
            Department: [],
            Nmb_opt: "",
            Duration: [],
            DurationQuantity: "",
            Work_Arr: [],
            LocationEn: "",
            LocationFr: "",
            sec_lvl: [],
            Language: [],
            LanguageComprehension: "",
            NoOpt: true,
            ContactEmail: "",
            ContactName: "",
            OptId: 0,
            Expired: false,
            pageLoading: true,
            deleteLoading: false,
            deleted: false,
            modalOpen: false,
            skills: [],
            WorkSchedule: []
        }

        if (!this.envValid()) 
            console.error('Check your env settings, something is missing!');
    }

    public async componentDidMount(): Promise<void> {
        await this._geturlID();
    }

    public _geturlID = async (): Promise<void> => {
        const params = new URLSearchParams(window.location.search);
        let val: number;

        val = params.get('JobOpportunityId') as unknown as number; // convert to number

        if (val !== null && val) { // check if value exist and not empty
            sessionStorage.setItem("JobOpportunityId", val.toString());
        } else {
            const sessionVal =  sessionStorage.getItem("JobOpportunityId");

            if (sessionVal !== null && sessionVal) {
                val = sessionVal as unknown as number;
            }
        }

        console.log("_geturlID val", val);

        if (val !== null && val) { // check if value exist and not empty
            this.setState({
                NoOpt: false,
                OptId: val
            })
            await this._getdetailsopt(val);
        } else {
            this.setState({
                NoOpt: true,
                pageLoading: false
            })
        }
    }
    
    public _getdetailsopt = async (valueid: number): Promise<void> => {
        const _sp: SPFI = getSP(this.props.context);

        try {
            const item = await _sp.web.lists.getByTitle("JobOpportunity").items.getById(valueid)
            .select(
                "Department", 
                "Department/NameEn", 
                "Department/NameFr", 
                "ClassificationCode", 
                "ClassificationCode/NameEn", 
                "ClassificationCode/NameFr", 
                "ClassificationLevel",
                "ClassificationLevel/NameEn",
                "ClassificationLevel/NameFr",
                "NumberOfOpportunities", 
                "JobTitleFr", 
                "JobTitleEn", 
                "JobDescriptionEn", 
                "JobDescriptionFr", 
                "ApplicationDeadlineDate", 
                "ContactEmail", 
                "ContactName", 
                "ProgramArea", 
                "JobType", 
                "Duration", 
                "Duration/NameEn", 
                "Duration/NameFr", 
                "DurationQuantity", 
                "WorkArrangement", 
                "WorkArrangement/NameEn", 
                "WorkArrangement/NameFr", 
                "WorkSchedule",
                "WorkSchedule/NameEn",
                "WorkSchedule/NameFr",
                "City",
                "City/Id", 
                "City/NameEn", 
                "City/NameFr",
                "SecurityClearance", 
                "SecurityClearance/NameEn", 
                "SecurityClearance/NameFr", 
                "LanguageRequirement", 
                "LanguageRequirement/NameEn", 
                "LanguageRequirement/NameFr", 
                "LanguageComprehension")
            .expand(
                "Department", 
                "ClassificationCode", 
                "ClassificationLevel",
                "Duration", 
                "WorkArrangement",
                "WorkSchedule",
                "City",
                "SecurityClearance", 
                "LanguageRequirement"
            )();
            console.log(item);

            const city = await _sp.web.lists.getByTitle("City").items.getById(item.City.Id)
            .select(
                "NameEn", 
                "NameFr", 
                "Region", 
                "Region/Id"
            )
            .expand(
                "Region"
            )();

            const region = await _sp.web.lists.getByTitle("Region").items.getById(city.Region.Id)
            .select(
                "NameEn", 
                "NameFr", 
                "Province", 
                "Province/Id",
                "Province/NameEn",
                "Province/NameFr"
            )
            .expand(
                "Province"
            )();

            const querySkills = await _sp.web.lists.getByTitle("JobOpportunity").items.getById(valueid)
            .select(
                "Skills/Id"
            )
            .expand(
                "Skills"
            )();

            const skillsEn: string[] = [];
            const skillsFr: string[] = [];
            if (querySkills && querySkills.Skills && querySkills.Skills.length > 0) {
                for (let i = 0; i < querySkills.Skills.length; i++) {
                    let skill = await _sp.web.lists.getByTitle("Skills").items.getById(querySkills.Skills[i].Id)();
                    
                    // Because our original list in dev was renamed we have to check for Title/Name...
                    skillsEn.push(skill.TitleEN ? skill.TitleEN : skill.NameEn);
                    skillsFr.push(skill.TitleFr ? skill.TitleFr : skill.NameFr);
                }
            }

            const expired = new Date() >= new Date(item.ApplicationDeadlineDate);
            
            this.setState({
                TitleFr: item.JobTitleFr,
                TitleEn: item.JobTitleEn,
                DescEn: item.JobDescriptionEn,
                DescFr: item.JobDescriptionFr,
                JobType: await this._get_terms(this.env.jobTypeTermSetId, item.JobType[0].TermGuid),
                program: await this._get_terms(this.env.programAreaTermSetId, item.ProgramArea[0].TermGuid),
                classification: `${item.ClassificationCode.NameEn}-${item.ClassificationLevel.NameEn}`,
                Department: item.Department,
                AppDeadline: item.ApplicationDeadlineDate.split('T')[0], // convert into format YYYY/MM/DD
                Nmb_opt: item.NumberOfOpportunities,
                Duration: item.Duration,
                DurationQuantity: item.DurationQuantity,
                Work_Arr:item.WorkArrangement,
                LocationEn: `${item.City.NameEn}, ${region.NameEn}, ${region.Province.NameEn}`,
                LocationFr: `${item.City.NameFr}, ${region.NameFr}, ${region.Province.NameFr}`,
                sec_lvl: item.SecurityClearance,
                Language: item.LanguageRequirement,
                ContactEmail: item.ContactEmail,
                ContactName: item.ContactName,
                LanguageComprehension: item.LanguageComprehension,
                Expired: expired,
                pageLoading: false,
                skills: {
                    en: skillsEn.join(', '),
                    fr: skillsFr.join(', ')
                },
                WorkSchedule: item.WorkSchedule
            })
        } catch(e) {
            console.error(e);
            this.setState({
                NoOpt: true,
                pageLoading: false
            });
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

        const info: TermStore.Term = await graph.termStore.groups.getById(this.env.careerMarketplaceTermSetId).sets.getById(termsetid).getTermById(termsid)();
        return JSON.parse(JSON.stringify(info.labels))[lang_id].name;
    }

    private toggleModal = (): void => {
        this.setState({modalOpen: !this.state.modalOpen});
    }

    private deleteOpportunity = async (): Promise<void> => {
        this.setState({
            deleteLoading: true
        });

        try {
            const aadClient: AadHttpClient = await this.props.context.aadHttpClientFactory.getClient(this.env.authClientId);

            const postOptions: IHttpClientOptions = {
                headers: {
                    "Content-Type": "application/json"
                },
                body: `{ItemId: ${this.state.OptId.toString()}}`
            };

            const response: HttpClientResponse = await aadClient.post(
                this.env.deleteApiUrl,
                AadHttpClient.configurations.v1,
                postOptions
            );

            if (response.ok) {
                this.setState({deleteLoading: false, deleted: true});
            } else {
                this.setState({deleteLoading: false});
            }
        } catch (e) {
            console.error(e);
        } finally {
            this.setState({deleteLoading: false, modalOpen: false});
        }
    }

    private getDeletedSubText = (): string => {
        return DOMPurify.sanitize(this.strings.oppDeletedSubText.replace('{jobTitle}', this.props.prefLang === 'fr-fr' ? this.state.TitleFr : this.state.TitleEn));
    }

    private populateRecoveryEmail = (): string => {
        const template = this.props.prefLang === 'fr-fr' ?
        "Bonjour,\n\nVeuillez récupérer la possibilité d’emploi supprimée dans le Carrefour de carrière, intitulée {jobTitle}, que j’ai supprimée le {date}.\n\nMerci,\n\n{name}" :
        "Hello,\n\nPlease recover the deleted Career Marketplace opportunity titled {jobTitle}, which I deleted on {date}.\n\nThank you,\n\n{name}";
        const today = new Date();
        const nameSplit = this.props.userDisplayName.split(',');
        const name = nameSplit.length > 1 ? `${nameSplit[1]} ${nameSplit[0]}` : this.state.ContactName;

        return template
            .replace('{jobTitle}', this.props.prefLang === 'fr-fr' ? this.state.TitleFr : this.state.TitleEn)
            .replace('{date}', today.toLocaleDateString())
            .replace('{name}', name);
    }

    private populateApplicationEmail = (): string => {
        const template = this.props.prefLang === 'fr-fr' ?
        `Bonjour {contactName},\n\nJ’espère que vous allez bien. Mon nom est {userName} et l’offre d’emploi que vous avez publiée dans le Carrefour d’emploi sur GCÉchange m’intéresse. Vous trouverez ci joint mon curriculum vitæ.\n\nMes compétences semblent correspondre à vos besoins et j’aimerais en discuter avec vous.\nJe vous remercie de prendre le temps de considérer ma candidature.\n\nCordialement,\n{userName}` :
        `Hello {contactName},\n\nI hope this message finds you well. My name is {userName}, and I am interested in the career opportunity you posted on the GCXchange Career Marketplace. Please find my resumé attached for your review.\n\nI would appreciate the opportunity to discuss how my skills align with your needs.\nThank you for your time and consideration.\n\nBest regards,\n{userName}`;
        
        const conNameSplit = this.state.ContactName.split(',');
        const contactName = conNameSplit.length > 1 ? `${conNameSplit[1]} ${conNameSplit[0]}` : this.state.ContactName;

        const usrNameSplit = this.props.userDisplayName.split(',');
        const userName = usrNameSplit.length > 1 ? `${usrNameSplit[1]} ${usrNameSplit[0]}` : this.props.userDisplayName;

        return template
            .replace(/{userName}/g, userName)
            .replace('{contactName}', contactName);
    }
 
    public render(): React.ReactElement<ISpfxCmDetailsProps> {
        const {
            hasTeamsContext,
        } = this.props;

        return this.state.deleted ? (
            <section className={`${styles.spfxCmDetails} ${hasTeamsContext ? styles.teams : ''}`}>
                <div className={styles.deletedSection}>
                    <h2 id={`cm-deleted-${this.state.OptId}-title`}>
                        {this.strings.oppDeletedTitle}
                    </h2>
                    <p dangerouslySetInnerHTML={{__html: this.getDeletedSubText()}}/>
                    <div className={styles.deletedButtons}>
                        <DefaultButton
                            text={this.strings.contactUs}
                            href={`mailto:support-soutien@gcx-gce.gc.ca?subject=${this.strings.emailSubject}&body=${encodeURIComponent(this.populateRecoveryEmail())}`}
                            aria-describedby={`cm-deleted-${this.state.OptId}-title`}
                            aria-label={this.strings.contactUs}
                        />
                        <PrimaryButton
                            text={this.strings.cmHomePage}
                            href={this.env.careerMarketplaceHomePage}
                            aria-describedby={`cm-deleted-${this.state.OptId}-title`}
                            aria-label={this.strings.cmHomePage}
                        />
                    </div>
                </div>
            </section>) : 
            this.state.pageLoading ? (
            <section className={`${styles.spfxCmDetails} ${hasTeamsContext ? styles.teams : ''}`}>
                <h2>{this.strings.loading}</h2>
            </section>) :
            (
            <section className={`${styles.spfxCmDetails} ${hasTeamsContext ? styles.teams : ''}`}>
                {this.state.NoOpt ? (

                    <div className={styles.welcome}>
                        <h2>Sorry! This opportunity do not exist+{this.state.NoOpt}</h2>
                        <p>Please, try another one or reach out to our support team!</p>
                    </div>

                ) : ( 
            <>
                {this.state.Expired ? (
                    <div className={styles.expiredBanner}>
                        <p role="status" aria-live="polite">
                            <Icon iconName='ChromeClose' /> &nbsp; {this.strings.Expired}
                        </p>
                    </div>
                ) : null}

                <div className={styles.retention}>
                    <p>
                        <span id="retention">
                            {parse(this.strings.Retention)}
                        </span>
                    </p>
                </div>

                <div className={styles.welcome}>
                    <h2>
                        <span id="JobTitle">{this.props.prefLang === "fr-fr" ? (this.state.TitleFr) : (this.state.TitleEn)}</span>
                    </h2>
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
                            <span className={styles.jobtype_space}>{this.strings.JobType} ({this.state.JobType})</span>
                            <span>{this.strings.ApplicationDeadline}: {this.state.AppDeadline}</span>
                        </div>
                        <div>
                            <h3>{this.strings.OpportunityDetails}:</h3>
                            <p>
                                <h4>{this.strings.ProgramArea}</h4>
                                {this.state.program}
                            </p>
                            <p>
                                <h4>{this.strings.classification}</h4>
                                {this.state.classification}
                            </p>
                            <p>
                                <h4>{this.strings.Department}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.Department.NameFr) : (this.state.Department.NameEn)}
                            </p>
                            <p>
                                <h4>{strings.NumberOpportunities}</h4>
                                {this.state.Nmb_opt}
                            </p>
                            <p>
                                <h4>{this.strings.Duration}</h4>
                                {this.state.DurationQuantity + " "}
                                {this.props.prefLang === "fr-fr" ? (this.state.Duration.NameFr) : (this.state.Duration.NameEn)}
                            </p>
                            <p>
                                <h4>{this.strings.ApplicationDeadline}</h4>
                                {this.state.AppDeadline}
                            </p>
                            <p>
                                <h4>{this.strings.WorkArrangement}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.Work_Arr.NameFr) : (this.state.Work_Arr.NameEn)}
                            </p>
                            <p>
                                <h4>{this.strings.Location}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.LocationFr) : (this.state.LocationEn)}
                            </p>
                            <p>
                                <h4>{this.strings.SecurityLevel}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.sec_lvl.NameFr) : (this.state.sec_lvl.NameEn)}
                            </p>
                            <p>
                                <h4>{this.strings.LanguageRequirements}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.Language.NameFr) : (this.state.Language.NameEn)}
                                {' '}{ this.state.LanguageComprehension }
                            </p>
                            <p>
                                <h4>{this.strings.skills}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.skills.fr) : (this.state.skills.en)}
                            </p>
                            <p>
                                <h4>{this.strings.workSchedule}</h4>
                                {this.props.prefLang === "fr-fr" ? (this.state.WorkSchedule.NameFr) : (this.state.WorkSchedule.NameEn)}
                            </p>
                        </div>
                        {this.props.prefLang === "fr-fr" ? (
                            <PrimaryButton 
                                text={this.state.Expired ? this.strings.ApplicationsClosed : this.strings.Apply} 
                                disabled={this.state.Expired} 
                                styles={{rootDisabled: {backgroundColor: '#403F3F', color: '#FFF'}}} 
                                href={`mailto:${this.state.ContactEmail}?subject=${encodeURIComponent(`Intérêt pour l'opportunité ${this.state.TitleFr}`)}&body=${encodeURIComponent(this.populateApplicationEmail())}&JobOpportunityId=${this.state.OptId}`}
                                aria-describedby='JobTitle' 
                                aria-label={this.state.Expired ? this.strings.ApplicationsClosed : this.strings.Apply}
                            />
                        ) : (
                            <PrimaryButton 
                                text={this.state.Expired ? this.strings.ApplicationsClosed : this.strings.Apply} 
                                disabled={this.state.Expired} 
                                styles={{rootDisabled: {backgroundColor: '#403F3F', color: '#FFF'}}} 
                                href={`mailto:${this.state.ContactEmail}?subject=${encodeURIComponent(`Interested in the ${this.state.TitleEn} opportunity`)}&body=${encodeURIComponent(this.populateApplicationEmail())}&JobOpportunityId=${this.state.OptId}`}
                                aria-describedby='JobTitle'
                                aria-label={this.state.Expired ? this.strings.ApplicationsClosed : this.strings.Apply}
                            />
                        )}

                        {this.props.context.pageContext.user.email === this.state.ContactEmail ? (
                            <PrimaryButton 
                                className={styles.margin_edit_buttom} 
                                text={this.strings.Edit} 
                                onClick={() => {
                                    window.location.href = `${this.env.editOpportunityPage}${this.state.OptId}`
                                }}
                                aria-describedby='JobTitle'
                                aria-label={this.strings.Edit} 
                            />
                        ) : (<></>)}   

                        {this.props.context.pageContext.user.email === this.state.ContactEmail ? (
                            <PrimaryButton 
                                onClick={this.toggleModal} 
                                disabled={this.state.deleteLoading || this.state.deleted} 
                                className={styles.margin_edit_buttom} 
                                text={this.strings.Delete} 
                                styles={{ rootHovered: { backgroundColor: 'rgb(227 16 16)', borderColor: 'rgb(227 16 16)', color: '#FFF' }, root: { backgroundColor: '#A60404', borderColor: '#A60404', color: '#FFF' } }} 
                                aria-describedby='JobTitle'
                                aria-label={this.strings.Delete}
                                />
                        ) : (<></>)}  

                        <Modal 
                            isOpen={this.state.modalOpen} 
                            onDismiss={this.toggleModal}
                            styles={{main: {width: '50%', maxWidth: '585px', borderRadius: '5px'}}}
                        >
                            <div className={`${styles.deleteModal}`}>
                                <div className={`${styles.modalHeader}`}>
                                    <h2 id={`cm-delete-${this.state.OptId}-title`}>
                                        {this.strings.dialogTitle}
                                    </h2>
                                    <IconButton 
                                        onClick={this.toggleModal} 
                                        iconProps={{iconName: 'ChromeClose'}} 
                                        styles={{icon: {color: 'inherit', backgroundColor: 'transparent', fontSize: 'small'}}}
                                        aria-describedby={`cm-delete-${this.state.OptId}-title`}
                                        aria-label={this.strings.cancel}
                                    />
                                </div>
                                <p 
                                    dangerouslySetInnerHTML={{__html: DOMPurify.sanitize(this.strings.dialogText)}} 
                                    id={`cm-delete-${this.state.OptId}-content`} 
                                />
                                <div className={`${styles.modalActions}`}>
                                    <DefaultButton 
                                        text={this.strings.cancel}
                                        onClick={this.toggleModal}
                                        aria-describedby={`cm-delete-${this.state.OptId}-title`}
                                        aria-label={this.strings.cancel}
                                    />
                                    <PrimaryButton 
                                        onClick={this.deleteOpportunity} 
                                        disabled={this.state.deleteLoading || this.state.deleted} 
                                        text={this.strings.Delete} 
                                        styles={this.state.deleteLoading ? undefined : { rootHovered: { backgroundColor: 'rgb(227 16 16)', borderColor: 'rgb(227 16 16)', color: '#FFF' }, root: { backgroundColor: '#A60404', borderColor: '#A60404', color: '#FFF' } }} 
                                        aria-describedby={`cm-delete-${this.state.OptId}-title`}
                                        aria-label={this.strings.Delete}
                                    />
                                </div>
                            </div>
                        </Modal>  
                </div>
                </>    
                )
            }
        </section>
        );
    }
}
