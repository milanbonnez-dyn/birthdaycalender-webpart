import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BirthdayCalendarWebPartStrings';
import BirthdayCalendar from './components/BirthdayCalendar';
import { IBirthdayCalendarProps, IBirthdayCalendarWebPartProps, LeapYearStrategy } from './models/IBirthday';

/**
 * Birthday Calendar Web Part
 *
 * Displays upcoming birthdays from a SharePoint Events list with proper
 * yearly recurrence handling, including leap year support.
 */
export default class BirthdayCalendarWebPart extends BaseClientSideWebPart<IBirthdayCalendarWebPartProps> {
  private _isDarkTheme: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IBirthdayCalendarProps> = React.createElement(
      BirthdayCalendar,
      {
        eventsListTitle: this.properties.eventsListTitle || 'Verjaardagen',
        birthdayCategory: this.properties.birthdayCategory !== undefined ? this.properties.birthdayCategory : '',
        daysAhead: this.properties.daysAhead || 30,
        showAge: this.properties.showAge || false,
        leapYearStrategy: this.properties.leapYearStrategy || LeapYearStrategy.February28,
        maxItems: this.properties.maxItems || 10,
        webPartTitle: this.properties.webPartTitle || strings.DefaultTitle,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        isDarkTheme: this._isDarkTheme
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;

    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneTextField('eventsListTitle', {
                  label: strings.EventsListTitleLabel,
                  description: strings.EventsListTitleDescription,
                  value: this.properties.eventsListTitle || 'Verjaardagen'
                }),
                PropertyPaneTextField('birthdayCategory', {
                  label: strings.BirthdayCategoryLabel,
                  description: strings.BirthdayCategoryDescription,
                  value: this.properties.birthdayCategory || 'Birthday'
                })
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('webPartTitle', {
                  label: strings.WebPartTitleLabel,
                  value: this.properties.webPartTitle || strings.DefaultTitle
                }),
                PropertyPaneSlider('daysAhead', {
                  label: strings.DaysAheadLabel,
                  min: 7,
                  max: 365,
                  step: 7,
                  value: this.properties.daysAhead || 30,
                  showValue: true
                }),
                PropertyPaneSlider('maxItems', {
                  label: strings.MaxItemsLabel,
                  min: 1,
                  max: 50,
                  step: 1,
                  value: this.properties.maxItems || 10,
                  showValue: true
                }),
                PropertyPaneToggle('showAge', {
                  label: strings.ShowAgeLabel,
                  onText: strings.ShowAgeOnText,
                  offText: strings.ShowAgeOffText,
                  checked: this.properties.showAge || false
                })
              ]
            },
            {
              groupName: strings.LeapYearGroupName,
              groupFields: [
                PropertyPaneDropdown('leapYearStrategy', {
                  label: strings.LeapYearStrategyLabel,
                  options: [
                    { key: LeapYearStrategy.February28, text: strings.LeapYearFeb28 },
                    { key: LeapYearStrategy.March1, text: strings.LeapYearMar1 }
                  ],
                  selectedKey: this.properties.leapYearStrategy || LeapYearStrategy.February28
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
