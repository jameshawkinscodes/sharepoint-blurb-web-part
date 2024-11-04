// IBlurbProps.ts
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
  }>;
  containerCount: number;
  onContainerClick: (index: number) => void;
  onEditClick: (index: number) => void;
  onMoveClick: (index: number, direction: 'up' | 'down') => void; // Updated to include direction
  onRemoveClick: (index: number, updatedCount: number) => void; // Updated to include updated count
}
