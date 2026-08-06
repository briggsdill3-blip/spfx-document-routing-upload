import { SPFI } from '@pnp/sp';

export interface ISiteContentIndexProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  targetSites: string;
  includeSystemLists: boolean;
  groupBySite: boolean;
}