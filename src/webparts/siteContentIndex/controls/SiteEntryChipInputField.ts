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

interface IInternalProperties extends ISiteEntryChipInputFieldProps {
  onRender: (elem: HTMLElement) => void;
  onDispose: (elem: HTMLElement) => void;
}

class SiteEntryChipInputField implements IPropertyPaneField<IInternalProperties> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IInternalProperties;
  private _originalProps: ISiteEntryChipInputFieldProps;

  constructor(targetProperty: string, properties: ISiteEntryChipInputFieldProps) {
    this.targetProperty = targetProperty;
    this._originalProps = properties;
    this.properties = {
      ...properties,
      onRender: this._onRender.bind(this),
      onDispose: this._onDispose.bind(this)
    };
  }

  private _onRender(elem: HTMLElement): void {
    if (!elem) {
      return;
    }

    elem.innerHTML = '';

    const labelElem = document.createElement('div');
    labelElem.textContent = this._originalProps.label || 'Target Sites';
    labelElem.style.fontWeight = '600';
    labelElem.style.fontSize = '14px';
    labelElem.style.marginBottom = '6px';
    elem.appendChild(labelElem);

    const controlContainer = document.createElement('div');
    elem.appendChild(controlContainer);

    const element: React.ReactElement<ISiteEntryChipInputProps> = React.createElement(SiteEntryChipInput, {
      entries: this._originalProps.entries,
      onChange: (entries: ISiteEntry[]) => {
        this._originalProps.onChanged(this.targetProperty, entries);
      }
    });

    ReactDom.render(element, controlContainer);
  }

  private _onDispose(elem: HTMLElement): void {
    if (elem) {
      ReactDom.unmountComponentAtNode(elem);
    }
  }
}

export function PropertyPaneSiteEntryChipInput(
  targetProperty: string,
  properties: ISiteEntryChipInputFieldProps
): IPropertyPaneField<IInternalProperties> {
  return new SiteEntryChipInputField(targetProperty, properties);
}