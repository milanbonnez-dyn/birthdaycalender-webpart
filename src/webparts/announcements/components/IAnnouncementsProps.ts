import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IAnnouncementsProps {
  context: WebPartContext;
  siteUrl: string;
  announcementsListTitle: string;
  dismissedListTitle: string;
}
