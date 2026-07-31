import { SPFI } from '@pnp/sp';

export interface IRecognitionFeedProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  siteAbsoluteUrl: string;
  listId: string;
  fieldName: string;
  fieldPhotoEmail: string;
  fieldAwardType: string;
  fieldMonth: string;
  fieldYear: string;
  fieldJustification: string;
  fieldTeamFlag: string;
  fieldTeamName: string;
  fieldTeamMembers: string;
  monthsToShow: number;
  scrollSpeed: 'slow' | 'medium' | 'off';
  scrollDirection: 'left' | 'right';
  nominateUrl: string;
}