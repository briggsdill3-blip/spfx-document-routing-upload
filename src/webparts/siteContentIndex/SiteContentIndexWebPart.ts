import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { spfi, SPFx as spSPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

import * as strings from 'SiteContentIndexWebPartStrings';
import SiteContentIndex from './components/SiteContentIndex';
import { ISiteContentIndexProps } from './components/ISiteContentIndexProps';
import { ISiteEntry } from './components/ISiteEntry';
import { PropertyPaneSiteEntryChipInput } from './controls/SiteEntryChipInputField';

export interface ISiteContentIndexWebPartProps {
  description: string;
  targetSites: ISiteEntry[];
  includeSystemLists: boolean;
  groupBySite: boolean;
}

export default class SiteContentIndexWebPart extends BaseClientSideWebPart<ISiteContentIndexWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _environmentMessage: string = '';
  private _sp!: SPFI;

  public render(): void {
    if (!this._sp) {
      return;
    }

    const element: React.ReactElement<ISiteContentIndexProps> = React.createElement(
      SiteContentIndex,
      {
        description: this.properties.description,
        theme: this._theme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this._sp,
        targetSites: this.properties.targetSites || [],
        includeSystemLists: this.properties.includeSystemLists || false,
        groupBySite: this.properties.groupBySite !== undefined ? this.properties.groupBySite : true
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._sp = spfi().using(spSPFx(this.context));

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._theme = currentTheme;
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
              groupName: strings.ConfigurationGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSiteEntryChipInput('targetSites', {
                  label: strings.TargetSitesFieldLabel,
                  entries: this.properties.targetSites || [],
                  theme: this._theme,
                  onChanged: (targetProperty: string, entries: ISiteEntry[]) => {
                    this.properties.targetSites = entries;
                    this.render();
                  }
                }),
                PropertyPaneToggle('includeSystemLists', {
                  label: strings.IncludeSystemListsFieldLabel,
                  onText: 'Shown',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('groupBySite', {
                  label: strings.GroupBySiteFieldLabel,
                  onText: 'Grouped',
                  offText: 'Flat list'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}