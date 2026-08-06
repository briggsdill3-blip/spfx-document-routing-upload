import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import SiteEntryChipInput, { ISiteEntryChipInputProps } from './SiteEntryChipInput';
import { ISiteEntry } from '../components/ISiteEntry';

export interface ISiteEntryChipInputFieldProps {
  label: string;
  entries: ISiteEntry[];
  theme: IReadonlyTheme | undefined;
  onChanged: (targetProperty: string, entries: ISiteEntry[]) => void;
}

class SiteEntryChipInputField implements IPropertyPaneField<ISiteEntryChipInputFieldProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: ISiteEntryChipInputFieldProps;

  constructor(targetProperty: string, properties: ISiteEntryChipInputFieldProps) {
    this.targetProperty = targetProperty;
    this.properties = properties;
  }

  public render(elem: HTMLElement): void {
    if (!elem) {
      return;
    }

    const themeColors = this.properties.theme ? this.properties.theme.semanticColors : undefined;
    const themePalette = this.properties.theme ? this.properties.theme.palette : undefined;

    const bodyText: string = (themeColors && themeColors.bodyText) || '#F5F5F0';
    const bodyDivider: string = (themeColors && themeColors.bodyDivider) || (themePalette && themePalette.neutralLight) || '#3A3A3A';
    const bodyBackground: string = (themeColors && themeColors.bodyBackground) || '#1E1E1E';
    const themePrimary: string = (themePalette && themePalette.themePrimary) || '#BF9B30';
    const paletteWhite: string = (themePalette && themePalette.white) || '#FFFFFF';

    elem.style.setProperty('--scix-pane-text', bodyText);
    elem.style.setProperty('--scix-pane-border', bodyDivider);
    elem.style.setProperty('--scix-pane-input-bg', bodyBackground);
    elem.style.setProperty('--scix-pane-accent', themePrimary);
    elem.style.setProperty('--scix-pane-accent-text', paletteWhite);

    const labelElem = document.createElement('div');
    labelElem.textContent = this.properties.label || '';
    labelElem.style.fontWeight = '600';
    labelElem.style.fontSize = '14px';
    labelElem.style.marginBottom = '6px';
    labelElem.style.color = bodyText;

    const controlContainer = document.createElement('div');

    elem.innerHTML = '';
    elem.appendChild(labelElem);
    elem.appendChild(controlContainer);

    const element: React.ReactElement<ISiteEntryChipInputProps> = React.createElement(SiteEntryChipInput, {
      entries: this.properties.entries,
      onChange: (entries: ISiteEntry[]) => {
        this.properties.entries = entries;
        this.properties.onChanged(this.targetProperty, entries);
      }
    });

    ReactDom.render(element, controlContainer);
  }
}

export function PropertyPaneSiteEntryChipInput(
  targetProperty: string,
  properties: ISiteEntryChipInputFieldProps
): IPropertyPaneField<ISiteEntryChipInputFieldProps> {
  return new SiteEntryChipInputField(targetProperty, properties);
}