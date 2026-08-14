import { SPFI } from '@pnp/sp';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISiteEntry } from './ISiteEntry';

export interface IDocumentUploadRouterProps {
  theme: IReadonlyTheme | undefined;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  targetSites: ISiteEntry[];
  showTitle: boolean;
  customTitle: string;
  tileDescription: string;
  tileIconName: string;
  accentColorOverride: string;
  tileBackgroundColorOverride: string;
}