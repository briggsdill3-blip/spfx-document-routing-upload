import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  IPropertyPaneCustomFieldProps,
  PropertyPaneCustomField
} from '@microsoft/sp-property-pane';
import SiteEntryChipInput, { ISiteEntryChipInputProps } from './SiteEntryChipInput';
import { ISiteEntry } from '../components/ISiteEntry';

export interface ISiteEntryChipInputFieldProps {
  label: string;
  entries: ISiteEntry[];
}

export function PropertyPaneSiteEntryChipInput(
  targetProperty: string,
  props: ISiteEntryChipInputFieldProps
): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  return PropertyPaneCustomField({
    key: `${targetProperty}-siteEntryChipInput`,
    onRender: (elem: HTMLElement, context?: unknown, changeCallback?: (targetProperty?: string, newValue?: unknown) => void) => {
      elem.innerHTML = '';

      const labelElem = document.createElement('div');
      labelElem.textContent = props.label || 'Target Sites';
      labelElem.style.fontWeight = '600';
      labelElem.style.fontSize = '14px';
      labelElem.style.marginBottom = '6px';
      elem.appendChild(labelElem);

      const controlContainer = document.createElement('div');
      elem.appendChild(controlContainer);

      const element: React.ReactElement<ISiteEntryChipInputProps> = React.createElement(SiteEntryChipInput, {
        entries: props.entries,
        onChange: (entries: ISiteEntry[]) => {
          if (changeCallback) {
            changeCallback(targetProperty, entries);
          }
        }
      });

      ReactDom.render(element, controlContainer);
    },
    onDispose: (elem: HTMLElement) => {
      if (elem) {
        ReactDom.unmountComponentAtNode(elem);
      }
    }
  });
}