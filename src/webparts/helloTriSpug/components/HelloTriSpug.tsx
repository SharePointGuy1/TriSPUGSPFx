import * as React from 'react';
import styles from './HelloTriSpug.module.scss';
import { IHelloTriSpugProps } from './IHelloTriSpugProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloTriSpug extends React.Component<IHelloTriSpugProps, void> {
  public render(): React.ReactElement<IHelloTriSpugProps> {
    return (
      <div className={styles.helloTriSpug}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p className="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p className="ms-font-l ms-fontColor-white">{escape(this.props.description)}</p>
              <p className="ms-font-l ms-fontColor-white">Purpose: {escape(this.props.purpose)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <div id='spListContainer'></div>
      </div>
    );
  }
}
