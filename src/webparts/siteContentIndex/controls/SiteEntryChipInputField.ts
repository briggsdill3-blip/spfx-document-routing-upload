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

    elem.innerHTML = '';

    const labelElem = document.createElement('div');
    labelElem.textContent = this.properties.label || 'Target Sites';
    labelElem.style.fontWeight = '600';
    labelElem.style.fontSize = '14px';
    labelElem.style.marginBottom = '6px';
    elem.appendChild(labelElem);

    const controlContainer = document.createElement('div');
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