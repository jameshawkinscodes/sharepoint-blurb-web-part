export interface IBlurbProps {
    containerCount: number;
    description: string;
    containers: Array<{
      icon: string;
      backgroundColor: string;
      borderColor: string;
      title: string;
      text: string;
    }>;
    onContainerClick: (index: number) => void;
    isDarkTheme?: boolean;
    environmentMessage?: string;
    hasTeamsContext?: boolean;
    userDisplayName?: string;
  }
  