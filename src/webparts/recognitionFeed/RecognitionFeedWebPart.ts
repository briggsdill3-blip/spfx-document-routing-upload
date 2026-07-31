import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { spfi, SPFx as spSPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy, IColumnReturnProperty } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';

import * as strings from 'RecognitionFeedWebPartStrings';
import RecognitionFeed from './components/RecognitionFeed';
import { IRecognitionFeedProps } from './components/IRecognitionFeedProps';

export interface IRecognitionFeedWebPartProps {
  description: string;
  listId: string;
  fieldName: string;
  fieldPhotoEmail: string;
  fieldAwardType: string;
  fieldMonth: string;
  fieldYear: string;
  fieldJustification: string;
  fieldTeamFlag: string;
  fieldTeamName: string;
  fieldTeamMembers: string;
  monthsToShow: string;
  scrollSpeed: string;
  scrollDirection: string;
  nominateUrl: string;
}

const parseMonthsToShow = (raw: string): number => {
  const parsed = parseInt(raw, 10);
  if (isNaN(parsed) || parsed < 1) {
    return 5;
  }
  return parsed;
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
        siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        listId: this.properties.listId || '',
        fieldName: this.properties.fieldName || '',
        fieldPhotoEmail: this.properties.fieldPhotoEmail || '',
        fieldAwardType: this.properties.fieldAwardType || '',
        fieldMonth: this.properties.fieldMonth || '',
        fieldYear: this.properties.fieldYear || '',
        fieldJustification: this.properties.fieldJustification || '',
        fieldTeamFlag: this.properties.fieldTeamFlag || '',
        fieldTeamName: this.properties.fieldTeamName || '',
        fieldTeamMembers: this.properties.fieldTeamMembers || '',
        monthsToShow: parseMonthsToShow(this.properties.monthsToShow),
        scrollSpeed: (this.properties.scrollSpeed as 'slow' | 'medium' | 'off') || 'slow',
        scrollDirection: (this.properties.scrollDirection as 'left' | 'right') || 'left',
        nominateUrl: this.properties.nominateUrl || ''
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
    this._isDarkTheme = !!currentTheme.isInverted;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Cast to `any` - @pnp/spfx-property-controls bundles its own internal copy of the
    // SPFx base context types, which can conflict with this project's own copy even
    // though the actual runtime object is identical. This is a known, documented
    // workaround for this package, not project-specific.
    const spfxContext = this.context as any;

    const columnPickerProps = {
      context: spfxContext,
      listId: this.properties.listId,
      disabled: !this.properties.listId,
      orderBy: PropertyFieldColumnPickerOrderBy.Title,
      onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
      properties: this.properties,
      deferredValidationTime: 0,
      multiSelect: false,
      displayHiddenColumns: false,
      columnReturnProperty: IColumnReturnProperty['Internal Name']
    };

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyFieldListPicker('listId', {
                  label: 'List',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: spfxContext,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                })
              ]
            },
            {
              groupName: 'Field Mapping',
              groupFields: [
                PropertyFieldColumnPicker('fieldName', { ...columnPickerProps, label: 'Name field', selectedColumn: this.properties.fieldName, key: 'fieldNameId' }),
                PropertyFieldColumnPicker('fieldPhotoEmail', { ...columnPickerProps, label: 'Photo/Email field', selectedColumn: this.properties.fieldPhotoEmail, key: 'fieldPhotoEmailId' }),
                PropertyFieldColumnPicker('fieldAwardType', { ...columnPickerProps, label: 'Award Type field', selectedColumn: this.properties.fieldAwardType, key: 'fieldAwardTypeId' }),
                PropertyFieldColumnPicker('fieldMonth', { ...columnPickerProps, label: 'Month field', selectedColumn: this.properties.fieldMonth, key: 'fieldMonthId' }),
                PropertyFieldColumnPicker('fieldYear', { ...columnPickerProps, label: 'Year field', selectedColumn: this.properties.fieldYear, key: 'fieldYearId' }),
                PropertyFieldColumnPicker('fieldJustification', { ...columnPickerProps, label: 'Justification field', selectedColumn: this.properties.fieldJustification, key: 'fieldJustificationId' }),
                PropertyFieldColumnPicker('fieldTeamFlag', { ...columnPickerProps, label: 'Team Award flag field', selectedColumn: this.properties.fieldTeamFlag, key: 'fieldTeamFlagId' }),
                PropertyFieldColumnPicker('fieldTeamName', { ...columnPickerProps, label: 'Team Name field', selectedColumn: this.properties.fieldTeamName, key: 'fieldTeamNameId' }),
                PropertyFieldColumnPicker('fieldTeamMembers', { ...columnPickerProps, label: 'Team Members field', selectedColumn: this.properties.fieldTeamMembers, key: 'fieldTeamMembersId' })
              ]
            },
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneTextField('monthsToShow', {
                  label: 'Number of months to show',
                  description: 'Default 5'
                }),
                PropertyPaneDropdown('scrollSpeed', {
                  label: 'Auto-scroll speed',
                  options: [
                    { key: 'slow', text: 'Slow' },
                    { key: 'medium', text: 'Medium' },
                    { key: 'off', text: 'Off' }
                  ],
                  selectedKey: this.properties.scrollSpeed || 'slow'
                }),
                PropertyPaneDropdown('scrollDirection', {
                  label: 'Scroll direction',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'right', text: 'Right' }
                  ],
                  selectedKey: this.properties.scrollDirection || 'left'
                })
              ]
            },
            {
              groupName: 'Nominate',
              groupFields: [
                PropertyPaneTextField('nominateUrl', {
                  label: 'Nomination form URL'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}