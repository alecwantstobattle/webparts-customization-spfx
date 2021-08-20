import * as React from 'react';
import { ProjectsProps } from './ProjectsProps';
import { ProjectsState } from './ProjectsState'

export class Projects extends React.Component<ProjectsProps, ProjectsState> {
    constructor(props: ProjectsProps, state: ProjectsState) {
        super(props);
        this.state = {
            items: []
        };
    }

    public getItems() {
        
    }

    public componentDidMount() {
        this.getItems();
    }

    public render(): React.ReactElement<ProjectsProps> {
        return <div>
            Hello from Projects Component
        </div>
    }
}