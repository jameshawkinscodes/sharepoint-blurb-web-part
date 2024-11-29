import { DisplayMode } from '@microsoft/sp-core-library';

export interface IBlurbProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  containers: Array<{
    fontColor: string;
    icon: string;
    backgroundColor: string;
    borderColor: string;
    borderRadius: string;
    title: string;
    text: string;
    linkUrl?: string; // New property for the clickable link in each container
    linkTarget?: "_self" | "_blank" | string;
  }>;
  containerCount: number;
  isEditMode: boolean;
  displayMode: DisplayMode;
  onContainerClick: (index: number) => void;
  onEditClick: (index: number) => void;
  onMoveClick: (index: number, direction: 'up' | 'down') => void; // Updated to include direction
  onRemoveClick: (index: number, updatedCount: number) => void; // Updated to include updated count
}
