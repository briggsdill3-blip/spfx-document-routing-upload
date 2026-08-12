import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';

export interface IColorPickerFieldProps {
  label: string;
  value: string;
  defaultColor: string;
  onPropertyChange: (propertyPath: string, newValue: string) => void;
}

interface IInternalProperties extends IColorPickerFieldProps {
  onRender?: (elem: HTMLElement) => void;
  onDispose?: (elem: HTMLElement) => void;
}

interface IColorSwatchControlProps {
  value: string;
  defaultColor: string;
  onChange: (newValue: string) => void;
}

const isValidHex = (raw: string): boolean => /^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$/.test(raw);

const ColorSwatchControl: React.FunctionComponent<IColorSwatchControlProps> = (props) => {
  const [hexInput, setHexInput] = React.useState<string>(props.value || '');

  React.useEffect(() => {
    setHexInput(props.value || '');
  }, [props.value]);

  const commitValue = (raw: string): void => {
    const trimmed = raw.trim();
    if (trimmed.length === 0) {
      props.onChange('');
      return;
    }
    const withHash = trimmed.charAt(0) === '#' ? trimmed : `#${trimmed}`;
    if (isValidHex(withHash)) {
      props.onChange(withHash);
    }
  };

  const swatchValue = isValidHex(props.value) ? props.value : props.defaultColor;

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
      <input
        type="color"
        value={swatchValue}
        onChange={(e) => {
          setHexInput(e.target.value);
          commitValue(e.target.value);
        }}
        style={{
          width: '32px',
          height: '28px',
          padding: 0,
          border: '1px solid #605E5C',
          borderRadius: '4px',
          cursor: 'pointer',
          background: 'none'
        }}
      />
      <input
        type="text"
        value={hexInput}
        placeholder={props.defaultColor}
        onChange={(e) => setHexInput(e.target.value)}
        onBlur={(e) => commitValue(e.target.value)}
        style={{
          flex: '1 1 auto',
          minWidth: 0,
          fontSize: '12px',
          padding: '4px 6px',
          border: '1px solid #605E5C',
          borderRadius: '4px'
        }}
      />
      {props.value && props.value.trim().length > 0 && (
        <button
          type="button"
          onClick={() => {
            setHexInput('');
            props.onChange('');
          }}
          style={{
            fontSize: '11px',
            background: 'none',
            border: 'none',
            color: '#0078D4',
            cursor: 'pointer',
            padding: '2px 4px'
          }}
        >
          Reset
        </button>
      )}
    </div>
  );
};

class ColorPickerField implements IPropertyPaneField<IInternalProperties> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public key: string;
  public properties: IInternalProperties;
  private _originalProps: IColorPickerFieldProps;

  constructor(targetProperty: string, properties: IColorPickerFieldProps) {
    this.targetProperty = targetProperty;
    this.key = targetProperty;
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
    labelElem.textContent = this._originalProps.label;
    labelElem.style.fontWeight = '600';
    labelElem.style.fontSize = '14px';
    labelElem.style.marginBottom = '6px';
    elem.appendChild(labelElem);

    const controlContainer = document.createElement('div');
    elem.appendChild(controlContainer);

    const element: React.ReactElement<IColorSwatchControlProps> = React.createElement(ColorSwatchControl, {
      value: this._originalProps.value,
      defaultColor: this._originalProps.defaultColor,
      onChange: (newValue: string) => {
        this._originalProps.value = newValue;
        this._originalProps.onPropertyChange(this.targetProperty, newValue);
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

export function PropertyPaneColorPickerField(
  targetProperty: string,
  properties: IColorPickerFieldProps
): IPropertyPaneField<IInternalProperties> {
  return new ColorPickerField(targetProperty, properties);
}