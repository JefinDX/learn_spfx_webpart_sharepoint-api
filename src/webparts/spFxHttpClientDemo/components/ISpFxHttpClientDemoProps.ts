import {
  ButtonClickedCallback,
  ICountryListItem
} from '../../../models';

export interface ISpFxHttpClientDemoProps {
  spListItems: ICountryListItem[];
  onGetListItems: ButtonClickedCallback;
  onAddListItem: ButtonClickedCallback;
  onUpdateListItem: ButtonClickedCallback;
  onDeleteListItem: ButtonClickedCallback;
  onUploadBtnClick: (fileData: ArrayBuffer, fileName: string) => Promise<void>;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

