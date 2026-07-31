import { SPFI } from '@pnp/sp';

export interface IRecognitionFeedProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  sp: SPFI;
  lockedLibrary: string;
  goodThreshold: number;
  warnThreshold: number;
  excludedFields: string[];
}