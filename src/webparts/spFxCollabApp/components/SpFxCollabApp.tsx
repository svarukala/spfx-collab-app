import * as React from 'react';
import styles from './SpFxCollabApp.module.scss';
import { ISpFxCollabAppProps } from './ISpFxCollabAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import AppInit from './AppInit';

export default class SpFxCollabApp extends React.Component<ISpFxCollabAppProps, {}> {
  public render(): React.ReactElement<ISpFxCollabAppProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      spoContext
    } = this.props;

    return (
      <section className={`${styles.spFxCollabApp} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Welcome, {escape(userDisplayName)}!</h2>
          <div>Web part property value: <strong>{escape(spoContext.pageContext.user.loginName)}</strong></div>
        </div>
        <div>
          <AppInit spoContext={spoContext} />
        </div>
      </section>
    );
  }
}
