import * as React from 'react';
import styles from './ProjectDashboard.module.scss';
import { IProjectDashboardProps } from './IProjectDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Projects } from './Projects/Projects';

export default class ProjectDashboard extends React.Component<IProjectDashboardProps, {}> {
  public render(): React.ReactElement<IProjectDashboardProps> {
    return (
      <div className={ styles.projectDashboard }>
        <Projects context= {this.props.context}></Projects>
      </div>
    );
  }
}
