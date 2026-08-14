import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import { spfi, SPFx as spSPFx, SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/site-users/web';

import * as strings from 'DocumentUploadRouterWebPartStrings';
import DocumentUploadRouter from './components/DocumentUploadRouter';
import { IDocumentUploadRouterProps } from './components/IDocumentUploadRouterProps';
import { ISiteEntry } from './components/ISiteEntry';
import { PropertyPaneSiteEntryChipInputField } from './controls/SiteEntryChipInputField';

export interface IDocumentUploadRouterWebPartProps {
  targetSites: ISiteEntry[];
  showTitle: boolean;
  customTitle: string;
  tileDescription: string;
  tileIconName: string;
  useThemeColors: boolean;
  accentColorOverride: string;
  tileBackgroundColorOverride: string;
  panelBackgroundColorOverride: string;
  panelTextColorOverride: string;
  panelBorderColorOverride: string;
}

const FALLBACK_ACCENT = '#BF9B30';
const FALLBACK_TILE_BG = '#1E1E1E';
const FALLBACK_PANEL_BG = '#1E1E1E';
const FALLBACK_TEXT = '#F5F5F0';
const FALLBACK_BORDER = '#3A3A3A';

const ICON_OPTIONS = [
  { key: 'CloudUpload', text: 'Cloud upload' },
  { key: 'Upload', text: 'Upload' },
  { key: 'Page', text: 'Page' },
  { key: 'Documentation', text: 'Documentation' },
  { key: 'FabricFolder', text: 'Folder' },
  { key: 'Archive', text: 'Archive' },
  { key: 'Send', text: 'Send' },
  { key: 'Save', text: 'Save' },
  { key: 'Attach', text: 'Attach' },
  { key: 'DocumentSet', text: 'Document set' }
];

export default class DocumentUploadRouterWebPart extends BaseClientSideWebPart<IDocumentUploadRouterWebPartProps> {

  private _theme: IReadonlyTheme | undefined;
  private _environmentMessage: string = '';
  private _sp!: SPFI;
  private _targetSitesField: PropertyPaneSiteEntryChipInputField | undefined;

  public render(): void {
    if (!this._sp) {
      return;
    }

    const element: React.ReactElement<IDocumentUploadRouterProps> = React.createElement(
      DocumentUploadRouter,
      {
        theme: this._theme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this._sp,
        webPartContext: this.context,
        targetSites: this.properties.targetSites || [],
        showTitle: this.properties.showTitle !== undefined ? this.properties.showTitle : true,
        customTitle: this.properties.customTitle || '',
        tileDescription: this.properties.tileDescription || '',
        tileIconName: this.properties.tileIconName || 'CloudUpload',
        useThemeColors: this.properties.useThemeColors !== undefined ? this.properties.useThemeColors : true,
        accentColorOverride: this.properties.accentColorOverride || '',
        tileBackgroundColorOverride: this.properties.tileBackgroundColorOverride || '',
        panelBackgroundColorOverride: this.properties.panelBackgroundColorOverride || '',
        panelTextColorOverride: this.properties.panelTextColorOverride || '',
        panelBorderColorOverride: this.properties.panelBorderColorOverride || ''
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    initializeIcons();
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
    return Version.parse('3.1');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'useThemeColors') {
      this.context.propertyPane.refresh();
    }
    this.render();
  }

  private _handleColorFieldChange(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    (this.properties as unknown as Record<string, unknown>)[propertyPath] = newValue;
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const themePalette = this._theme ? this._theme.palette : undefined;
    const themeColors = this._theme ? this._theme.semanticColors : undefined;
    const accentDefault: string = (themePalette && themePalette.themePrimary) || FALLBACK_ACCENT;
    const tileBgDefault: string = (themeColors && themeColors.bodyBackground) || FALLBACK_TILE_BG;
    const panelBgDefault: string = (themeColors && themeColors.bodyBackground) || FALLBACK_PANEL_BG;
    const textDefault: string = (themeColors && themeColors.bodyText) || FALLBACK_TEXT;
    const borderDefault: string = (themeColors && themeColors.bodyDivider) || FALLBACK_BORDER;

    this._targetSitesField = new PropertyPaneSiteEntryChipInputField('targetSites', {
      label: strings.TargetSitesFieldLabel,
      entries: this.properties.targetSites || [],
      sp: this._sp,
      accentColor: accentDefault,
      onPropertyChange: (propertyPath: string, newValue: ISiteEntry[]) => {
        this.properties.targetSites = newValue;
        this.render();
        if (this._targetSitesField) {
          this._targetSitesField.properties.entries = newValue;
          this._targetSitesField.render();
        }
      }
    });

    const useThemeColors = this.properties.useThemeColors !== undefined ? this.properties.useThemeColors : true;

    const appearanceFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle('useThemeColors', {
        label: 'Colors',
        onText: 'Automatic',
        offText: 'Custom'
      })
    ];

    if (!useThemeColors) {
      appearanceFields.push(
        PropertyFieldColorPicker('accentColorOverride', {
          label: 'Accent Color',
          selectedColor: this.properties.accentColorOverride || accentDefault,
          onPropertyChange: this._handleColorFieldChange.bind(this),
          properties: this.properties,
          key: 'accentColorOverrideField',
          style: PropertyFieldColorPickerStyle.Full
        }) as IPropertyPaneField<unknown>,
        PropertyFieldColorPicker('tileBackgroundColorOverride', {
          label: 'Floating Button Background',
          selectedColor: this.properties.tileBackgroundColorOverride || tileBgDefault,
          onPropertyChange: this._handleColorFieldChange.bind(this),
          properties: this.properties,
          key: 'tileBackgroundColorOverrideField',
          style: PropertyFieldColorPickerStyle.Full
        }) as IPropertyPaneField<unknown>,
        PropertyFieldColorPicker('panelBackgroundColorOverride', {
          label: 'Panel Background',
          selectedColor: this.properties.panelBackgroundColorOverride || panelBgDefault,
          onPropertyChange: this._handleColorFieldChange.bind(this),
          properties: this.properties,
          key: 'panelBackgroundColorOverrideField',
          style: PropertyFieldColorPickerStyle.Full
        }) as IPropertyPaneField<unknown>,
        PropertyFieldColorPicker('panelTextColorOverride', {
          label: 'Panel Text',
          selectedColor: this.properties.panelTextColorOverride || textDefault,
          onPropertyChange: this._handleColorFieldChange.bind(this),
          properties: this.properties,
          key: 'panelTextColorOverrideField',
          style: PropertyFieldColorPickerStyle.Full
        }) as IPropertyPaneField<unknown>,
        PropertyFieldColorPicker('panelBorderColorOverride', {
          label: 'Panel Borders',
          selectedColor: this.properties.panelBorderColorOverride || borderDefault,
          onPropertyChange: this._handleColorFieldChange.bind(this),
          properties: this.properties,
          key: 'panelBorderColorOverrideField',
          style: PropertyFieldColorPickerStyle.Full
        }) as IPropertyPaneField<unknown>
      );
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.TitleGroupName,
              groupFields: [
                PropertyPaneToggle('showTitle', {
                  label: strings.ShowTitleFieldLabel,
                  onText: 'Shown',
                  offText: 'Hidden'
                }),
                PropertyPaneTextField('customTitle', {
                  label: strings.CustomTitleFieldLabel,
                  description: 'Leave blank to use the default: Document Upload Router',
                  disabled: !this.properties.showTitle
                }),
                PropertyPaneTextField('tileDescription', {
                  label: 'Tile Description',
                  description: 'Short text shown on the button before it opens.'
                }),
                PropertyPaneDropdown('tileIconName', {
                  label: 'Tile Icon',
                  options: ICON_OPTIONS,
                  selectedKey: this.properties.tileIconName || 'CloudUpload'
                })
              ]
            },
            {
              groupName: strings.ConfigurationGroupName,
              groupFields: [
                this._targetSitesField
              ]
            },
            {
              groupName: 'Appearance',
              groupFields: appearanceFields
            }
          ]
        }
      ]
    };
  }
}