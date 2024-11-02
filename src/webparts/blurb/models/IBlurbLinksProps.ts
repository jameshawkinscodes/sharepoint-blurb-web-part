// IBlurbLinkProps.ts
import { DisplayMode } from "@microsoft/sp-core-library";
import { IBlurb } from "./IBlurb";  // Assuming IBlurb is also in models

export interface IBlurbLinkProps {
  webPartTitle: string;
  setWebpartTitle: (val: string) => void;
  links: IBlurb[];
  setLinks: (val: IBlurb[]) => void;
  SelectedItemId: string;
  setSelectedItemId: (id: string) => void;
  displayMode: DisplayMode;
}
