import { SPFI } from '@pnp/sp';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISiteEntry } from './ISiteEntry';

export interface IDocumentUploadRouterProps {
  theme: IReadonlyTheme | undefined;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  webPartContext: WebPartContext;
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