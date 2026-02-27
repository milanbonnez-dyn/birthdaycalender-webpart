import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import Announcements from './components/Announcements';
import { IAnnouncementsProps } from './components/IAnnouncementsProps';
import { sp } from "@pnp/sp";


export interface IAnnouncementsWebPartProps {
  // voorlopig geen properties nodig
}

export default class AnnouncementsWebPart
  extends BaseClientSideWebPart<IAnnouncementsWebPartProps> {

  /** ⬇⬇⬇ HIER hoort onInit ⬇⬇⬇ */
protected onInit(): Promise<void> {
  return super.onInit().then(_ => {
    sp.setup({
      spfxContext: this.context as any
    });
  });
}

  public render(): void {
    const element: React.ReactElement<IAnnouncementsProps> = React.createElement(
      Announcements,
      {
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        announcementsListTitle: 'Announcements',
        dismissedListTitle: 'Announcements_dismissed'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
