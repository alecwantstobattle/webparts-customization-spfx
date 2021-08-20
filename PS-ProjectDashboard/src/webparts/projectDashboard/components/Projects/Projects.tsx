import * as React from 'react';
import { ProjectsProps } from './ProjectsProps';
import { ProjectsState } from './ProjectsState';
import { 
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions
} from '@microsoft/sp-http';

export class Projects extends React.Component<ProjectsProps, ProjectsState> {
    constructor(props: ProjectsProps, state: ProjectsState) {
        super(props);
        this.state = {
            items: []
        };
    }

    public getItems() {
        this.props.context.spHttpClient
        .get(
            `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/Items?$expand=ProjectManager&$select=*,ProjectManager/EMail,ProjectManager/Title`,
            SPHttpClient.configurations.v1
        )
        .then(
            (response: SPHttpClientResponse): Promise<{ value: any[] }> => {
                return response.json();
            }
        )
        .then(
            (response: { value: any[] }) => {
                var _items = [];
                _items = _items.concat(response.value);
                this.setState({
                    items: _items,
                });
            }
        );
    }

    public componentDidMount() {
        this.getItems();
    }

    public render(): React.ReactElement<ProjectsProps> {
        return <div>
            <div>
                {this.state.items.map((item, key) => 
                    <li key={key}>
                        <h3>{item.Title}</h3>
                        <div>{item.ProjectManager.Title}</div>
                        <div>{item.Status}</div>
                    </li>
                )}
            </div>
        </div>;
    }
}