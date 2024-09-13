import * as React from 'react';
import styles from './SpFxHttpClientDemo.module.scss';
import type { ISpFxHttpClientDemoProps } from './ISpFxHttpClientDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpFxHttpClientDemo extends React.Component<ISpFxHttpClientDemoProps, {}> {

  private onGetListItemsClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    if (this.props.onGetListItems) this.props.onGetListItems();
  }

  private onAddListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    if (this.props.onAddListItem) this.props.onAddListItem();
  }

  private onUpdateListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    if (this.props.onUpdateListItem) this.props.onUpdateListItem();
  }

  private onDeleteListItemClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    if (this.props.onDeleteListItem) this.props.onDeleteListItem();
  }

  componentDidMount(): void {
    const inputFileElement = document.getElementsByClassName(`${styles.spFxHttpClientDemo}-fileUpload`)[0] as HTMLInputElement;
    const uploadButton = document.getElementsByClassName(`${styles.spFxHttpClientDemo}-uploadButton`)[0] as HTMLButtonElement;
    uploadButton.addEventListener('click', async () => {
      const filePathParts = inputFileElement.value.split('\\');
      const fileName = filePathParts[filePathParts.length - 1];
      if (inputFileElement.files) {
        const fileData = await this._getFileBuffer(inputFileElement.files[0]);
        await this.props.onUploadBtnClick(fileData, fileName);
      }
    });
  }

  private _getFileBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const fileReader = new FileReader();
      fileReader.onerror = (event: ProgressEvent<FileReader>) => {
        reject(event.target?.error);
      };
      fileReader.onloadend = (event: ProgressEvent<FileReader>) => {
        resolve(event.target?.result as ArrayBuffer);
      };
      fileReader.readAsArrayBuffer(file);
    });
  }



  public render(): React.ReactElement<ISpFxHttpClientDemoProps> {
    const {
      spListItems,
      // isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <>
        <section className={`${styles.spFxHttpClientDemo} ${hasTeamsContext ? styles.teams : ''}`}>
          <div className={styles.welcome}>
            {/* <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} /> */}
            <h2>Well done, {escape(userDisplayName)}!</h2>
            <div>{environmentMessage}</div>
          </div>
          <div className={styles.buttons}>
            <button type="button" onClick={this.onGetListItemsClicked}>Get Countries</button>
            <button type="button" onClick={this.onAddListItemClicked}>Add List Item</button>
            <button type="button" onClick={this.onUpdateListItemClicked}>Update List Item</button>
            <button type="button" onClick={this.onDeleteListItemClicked}>Delete List Item</button>
          </div>
          <div>
            <ul>
              {spListItems && spListItems.map((list) =>
                <li key={list.Id}>
                  <strong>Id:</strong> {list.Id}, <strong>Title:</strong> {list.Title}
                </li>
              )
              }
            </ul>
          </div>
        </section>
        <section className={`${styles.spFxHttpClientDemo} ${hasTeamsContext ? styles.teams : ''}`}>
          <div className={styles.inputs}>
            <input className={`${styles.spFxHttpClientDemo}-fileUpload`} type="file" /><br />
            <input className={`${styles.spFxHttpClientDemo}-uploadButton`} type="button" value="Upload" />
          </div>
        </section>
      </>
    );
  }

}
