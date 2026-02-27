import * as React from 'react';
import { IAnnouncementsProps } from './IAnnouncementsProps';
import { IAnnouncements } from './IAnnouncements';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './Announcements.module.scss';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { sp } from "@pnp/sp";



export interface IAnnouncementsState {
  loading: boolean;
  error?: string;
  announcements: IAnnouncements[];
}

export default class Announcements extends React.Component<IAnnouncementsProps, IAnnouncementsState> {

  constructor(props: IAnnouncementsProps) {
    super(props);

    this.state = {
      loading: true,
      announcements: []
    };
  }
  private getTypeClass(type?: string): string {
  switch (type) {
    case 'HR': return styles.typeHR;
    case 'Safety': return styles.typeSafety;
    case 'HSE': return styles.typeHSE;
    case 'Comm': return styles.typeComm;
    case 'A&C': return styles.typeAC;
    case 'Internal':
    default:
      return styles.typeInternal;
  }
}

  public componentDidMount(): void {
    this.loadAnnouncements();
  }

  private async loadAnnouncements(): Promise<void> {
    try {
      const currentUserEmail = this.props.context.pageContext.user.email;

      // 1. alle announcements ophalen
      const announcementsUrl =
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.announcementsListTitle}')/items?$select=Id,Title,Description,TypeAnn&$orderby=Created desc`;

      const announcementsResponse: SPHttpClientResponse =
        await this.props.context.spHttpClient.get(
          announcementsUrl,
          SPHttpClient.configurations.v1
        );

      const announcementsJson = await announcementsResponse.json();
      console.log("Announcements REST response:", announcementsJson);
      const allAnnouncements: IAnnouncements[] = announcementsJson.value;

      // 2. dismissed items voor huidige gebruiker
      // AANNAME: dc_announcements_dismissed heeft kolommen:
      // - AnnouncementId (Lookup naar dc_announcements) -> interne naam: AnnouncementId
      // - UserEmail (Single line of text) met e-mail van user
      //
      // PAS DIT AAN NAAR JOUW EIGEN KOLomnamen als dat anders is.
const dismissedUrl =
  `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.props.dismissedListTitle}')/items`
  + `?$select=Id,lookupAnnouncementID/Id,UserEmail/Id,UserEmail/EMail`
  + `&$expand=lookupAnnouncementID,UserEmail`
  + `&$filter=UserEmail/EMail eq '${encodeURIComponent(currentUserEmail)}'`;

      const dismissedResponse: SPHttpClientResponse =
        await this.props.context.spHttpClient.get(
          dismissedUrl,
          SPHttpClient.configurations.v1
        );

      const dismissedJson = await dismissedResponse.json();
       console.log("Dismissed Announcements REST response:", dismissedJson);
      const dismissedIds: number[] = dismissedJson.value.map((d: any) => d.lookupAnnouncementID.Id);


      // 3. filter: alleen items die NIET dismissed zijn door deze user
      const visibleAnnouncements = allAnnouncements.filter(a => dismissedIds.indexOf(a.Id) === -1);

      this.setState({
        loading: false,
        announcements: visibleAnnouncements,
        error: undefined
      });

    } catch (err: any) {
      this.setState({
        loading: false,
        error: err.message || 'Error loading announcements'
      });
    }
  }

private onDismiss = async (announcement: IAnnouncements): Promise<void> => {
  try {
    const user = await sp.web.currentUser.get();

    await sp.web.lists
      .getByTitle(this.props.dismissedListTitle)
      .items.add({
        Title: `Dismissed - ${announcement.Title}`,
        lookupAnnouncementIDId: announcement.Id,
        UserEmailId: user.Id, // 🔑 CRUCIAAL
        DismissDate: new Date()
      });

    this.setState(prev => ({
      announcements: prev.announcements.filter(a => a.Id !== announcement.Id)
    }));

  } catch (err) {
    console.error("Dismiss failed", err);
  }
};






  public render(): React.ReactElement<IAnnouncementsProps> {
    if (this.state.loading) {
      return <div>Loading announcements...</div>;
    }

    if (this.state.error) {
      return <div>Error: {this.state.error}</div>;
    }

    if (this.state.announcements.length === 0) {
      return <div>Geen nieuwe aankondigingen.</div>;
    }

return (
  <div className={styles.announcementsContainer}>
    {this.state.announcements.map(a => (
      <div
  key={a.Id}
  className={`${styles.card} ${this.getTypeClass(a.TypeAnn)}`}
>
        
        <div className={styles.cardHeader}>
          <h3 className={styles.cardTitle}>{a.Title}</h3>

          <button
  className={styles.dismissButton}
  title="Markeer als gelezen"
  aria-label="Markeer als gelezen"
  onClick={() => this.onDismiss(a)}
>
  ✓
</button>
        </div>

        {a.Description && (
          <p className={styles.cardBody}>{a.Description}</p>
        )}

      </div>
    ))}
  </div>
);

  }
}
