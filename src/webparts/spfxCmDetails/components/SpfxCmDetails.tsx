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
    Expired: boolean;
    pageLoading: boolean;
    deleteLoading: boolean;
    deleted: boolean;
    modalOpen: boolean;
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
            OptId: 0,
            Expired: false,
            pageLoading: true,
            deleteLoading: false,
            deleted: false,
            modalOpen: false
        }

        if (!this.envValid()) 
            console.error('Check your env settings, something is missing!');
    }

    public async componentDidMount(): Promise<void> {
        await this._geturlID();
    }

        //this works but gives me an inifite loop for can not find web
    // public async componentDidUpdate(prevPros: ISpfxCmDetailsProps, prevState: ISpfxCmDetailsState): Promise<void> {

    //     if (prevState.NoOpt !== this.state.NoOpt) {
    //         await this._geturlID();
    //     }
        
    // }
       

    public _geturlID = async (): Promise<void> => {
        const params = new URLSearchParams(window.location.search);
        const val = params.get('JobOpportunityId') as unknown as number; // convert to number
        if (val !== null && val ) { // check if value exist and not empty
            this.setState({
                NoOpt: false,
                OptId: val
            });
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

            const item = await _sp.web.lists.getByTitle("JobOpportunity").items.getById(valueid).select("Department", "Department/NameEn", "Department/NameFr", "ClassificationCode", "ClassificationCode/NameEn", "ClassificationCode/NameFr", "NumberOfOpportunities", "JobTitleFr", "JobTitleEn", "JobDescriptionEn", "JobDescriptionFr", "ApplicationDeadlineDate", "ContactEmail", "ProgramArea", "JobType", "Duration", "Duration/NameEn", "Duration/NameFr", "DurationQuantity", "WorkArrangement", "WorkArrangement/NameEn", "WorkArrangement/NameFr", "City", "City/NameEn", "City/NameFr", "SecurityClearance", "SecurityClearance/NameEn", "SecurityClearance/NameFr", "LanguageRequirement", "LanguageRequirement/NameEn", "LanguageRequirement/NameFr", "LanguageComprehension").expand("Department", "ClassificationCode", "Duration", "WorkArrangement", "City", "SecurityClearance", "LanguageRequirement")();
            console.log(item);

            const expired = new Date() >= new Date(item.ApplicationDeadlineDate);
            
            this.setState({
                TitleFr: item.JobTitleFr,
                TitleEn: item.JobTitleEn,
                DescEn: item.JobDescriptionEn,
                DescFr: item.JobDescriptionFr,
                JobType: await this._get_terms(this.env.jobTypeTermSetId, item.JobType[0].TermGuid),
                program: await this._get_terms(this.env.programAreaTermSetId, item.ProgramArea[0].TermGuid),
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
                LanguageComprehension: item.LanguageComprehension,
                Expired: expired,
                pageLoading: false
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
                            href={`mailto: support-soutien@gcx-gce.gc.ca`}
                            aria-labelledby={`cm-deleted-${this.state.OptId}-title`}
                        />
                        <PrimaryButton
                            text={this.strings.cmHomePage}
                            href={this.env.careerMarketplaceHomePage}
                            aria-labelledby={`cm-deleted-${this.state.OptId}-title`}
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
                                {this.props.prefLang === "fr-fr" ? (this.state.Location.NameFr) : (this.state.Location.NameEn)}
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
                        </div>
                        {this.props.prefLang === "fr-fr" ? (
                            <PrimaryButton 
                                text={this.state.Expired ? this.strings.ApplicationsClosed : this.strings.Apply} 
                                disabled={this.state.Expired} 
                                styles={{rootDisabled: {backgroundColor: '#403F3F', color: '#FFF'}}} 
                                href={`mailto: ${this.state.ContactEmail}?subject=Int%C3%A9r%C3%AAt%20envers%20une%20possibilit%C3%A9%20d%E2%80%99emploi&body=Le%20texte%20qui%20suit%20est%20un%20mod%C3%A8le%20de%20courriel.%20Vous%20n%E2%80%99avez%20qu%E2%80%99%C3%A0%20y%20ajouter%20les%20renseignements%20manquants%20(indiqu%C3%A9s%20en%20crochets)%20et%20%C3%A0%20modifier%20le%20texte%20si%20n%C3%A9cessaire.%5D%0A%0ABonjour%20%5Bnom%20sur%20l%E2%80%99offre%20d%E2%80%99emploi%5D%2C%0AJ%E2%80%99esp%C3%A8re%20que%20vous%20allez%20bien.%20Mon%20nom%20est%20%5Bvotre%20nom%5D%20et%20l%E2%80%99offre%20d%E2%80%99emploi%20que%20vous%20avez%20publi%C3%A9e%20dans%20le%20Carrefour%20d%E2%80%99emploi%20sur%20GC%C3%89change%20m%E2%80%99int%C3%A9resse.%20Vous%20trouverez%20ci%20joint%20mon%20curriculum%20vit%C3%A6.%0AMes%20comp%C3%A9tences%20semblent%20correspondre%20%C3%A0%20vos%20besoins%20et%20j%E2%80%99aimerais%20en%20discuter%20avec%20vous.%0AJe%20vous%20remercie%20de%20prendre%20le%20temps%20de%20consid%C3%A9rer%20ma%20candidature.%0ACordialement%2C%0A%5Bvotre%20nom%5D&JobOpportunityId=${this.state.OptId}`}
                                aria-labelledby='JobTitle' 
                            />
                        ) : (
                            <PrimaryButton 
                                text={this.state.Expired ? this.strings.ApplicationsClosed : this.strings.Apply} 
                                disabled={this.state.Expired} 
                                styles={{rootDisabled: {backgroundColor: '#403F3F', color: '#FFF'}}} 
                                href={`mailto: ${this.state.ContactEmail}?subject=Interested%20in%20Career%20Opportunity&body=%5BThe%20following%20is%20an%20email%20template.%20Simply%20fill%20in%20the%20missing%20information%20(indicated%20in%20brackets)%20and%20adjust%20the%20text%20as%20needed.%5D%0A%0AHello%20%5Bname%20on%20post%5D%2C%0AI%20hope%20this%20message%20finds%20you%20well.%20My%20name%20is%20%5Byour%20name%5D%2C%20and%20I%20am%20interested%20in%20the%20career%20opportunity%20you%20posted%20on%20the%20GCXchange%20Career%20Marketplace.%20Please%20find%20my%20resum%C3%A9%20attached%20for%20your%20review.%0AI%20would%20appreciate%20the%20opportunity%20to%20discuss%20how%20my%20skills%20align%20with%20your%20needs.%0AThank%20you%20for%20your%20time%20and%20consideration.%0ABest%20regards%2C%0A%5Byour%20name%5D%0A%0A&JobOpportunityId=${this.state.OptId}`}
                                aria-labelledby='JobTitle'
                            />
                        )}

                        {this.props.context.pageContext.user.email === this.state.ContactEmail ? (
                            <PrimaryButton 
                                className={styles.margin_edit_buttom} 
                                text={this.strings.Edit} 
                                href={`${this.env.editOpportunityPage}${this.state.OptId}`} 
                                aria-labelledby='JobTitle'
                            />
                        ) : (<></>)}   

                        {this.props.context.pageContext.user.email === this.state.ContactEmail ? (
                            <PrimaryButton 
                                onClick={this.toggleModal} 
                                disabled={this.state.deleteLoading || this.state.deleted} 
                                className={styles.margin_edit_buttom} text={this.strings.Delete} 
                                styles={{ rootHovered: { backgroundColor: 'rgb(227 16 16)', color: '#FFF' }, root: { backgroundColor: '#A60404', color: '#FFF' } }} 
                                aria-labelledby='JobTitle'
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
                                        aria-labelledby={`cm-delete-${this.state.OptId}-title`}
                                        ariaLabel={this.strings.cancel}
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
                                        aria-labelledby={`cm-delete-${this.state.OptId}-title`}
                                    />
                                    <PrimaryButton 
                                        onClick={this.deleteOpportunity} 
                                        disabled={this.state.deleteLoading || this.state.deleted} 
                                        text={this.strings.Delete} 
                                        styles={this.state.deleteLoading ? undefined : { rootHovered: { backgroundColor: 'rgb(227 16 16)', color: '#FFF' }, root: { backgroundColor: '#A60404', color: '#FFF' } }} 
                                        aria-labelledby={`cm-delete-${this.state.OptId}-title`}
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
