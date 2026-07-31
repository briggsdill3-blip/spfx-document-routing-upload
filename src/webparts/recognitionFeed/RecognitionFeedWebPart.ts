import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { spfi, SPFx as spSPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

import * as strings from 'RecognitionFeedWebPartStrings';
import RecognitionFeed from './components/RecognitionFeed';
import { IRecognitionFeedProps } from './components/IRecognitionFeedProps';

export interface IRecognitionFeedWebPartProps {
  description: string;
  lockedLibrary: string;
  goodThreshold: string;
  warnThreshold: string;
  excludedFields: string;
}

const parseThreshold = (raw: string, fallback: number): number => {
  const parsed = parseInt(raw, 10);
  if (isNaN(parsed) || parsed < 0 || parsed > 100) {
    return fallback;
  }
  return parsed;
};

const parseExcludedFields = (raw: string): string[] => {
  if (!raw) {
    return [];
  }
  return raw.split(',').map(f => f.trim().toLowerCase()).filter(f => f.length > 0);
};

export default class RecognitionFeedWebPart extends BaseClientSideWebPart<IRecognitionFeedWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp!: SPFI;

  public render(): void {
    const element: React.ReactElement<IRecognitionFeedProps> = React.createElement(
      RecognitionFeed,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this._sp,
        lockedLibrary: this.properties.lockedLibrary || '',
        goodThreshold: parseThreshold(this.properties.goodThreshold, 90),
        warnThreshold: parseThreshold(this.properties.warnThreshold, 70),
        excludedFields: parseExcludedFields(this.properties.excludedFields)
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
    if (!currentTheme || !this.domElement) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;

    const { semanticColors, palette } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);

      this.domElement.style.setProperty('--mcd-page-bg', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--mcd-card-bg', semanticColors.cardStandoutBackground || semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--mcd-text', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--mcd-text-secondary', semanticColors.bodySubtext || null);
      this.domElement.style.setProperty('--mcd-border', semanticColors.bodyDivider || null);
    }

    if (palette) {
      this.domElement.style.setProperty('--mcd-accent', palette.themePrimary || null);
      this.domElement.style.setProperty('--mcd-accent-dark', palette.themeDarkAlt || null);
      this.domElement.style.setProperty('--mcd-accent-light', palette.themeLighter || null);
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}