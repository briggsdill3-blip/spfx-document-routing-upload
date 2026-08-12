import { SPFI } from '@pnp/sp';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISiteEntry } from './ISiteEntry';

export interface ISiteContentIndexProps {
  theme: IReadonlyTheme | undefined;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  targetSites: ISiteEntry[];
  includeSystemLists: boolean;
  groupBySite: boolean;
  staleDaysThreshold: number;
  showTitle: boolean;
  customTitle: string;
  expandByDefault: boolean;
  enableRowStriping: boolean;
  accentColorOverride: string;
  stripeColorOverride: string;
}