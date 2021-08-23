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

    public getItems(filterVal) {
        if (filterVal === "*") {
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
        else {
            this.props.context.spHttpClient
            .get(
                `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/Items?$expand=ProjectManager&$select=*,ProjectManager/EMail,ProjectManager/Title&$filter=Status eq %27${filterVal}%27`,
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
    }

    public componentDidMount() {
        var getAll = "*";
        this.getItems(getAll);
    }

    public progFilter(filterVal) {
        switch(filterVal) {
            case "All":
                return this.getItems(filterVal);
            case "In Progress":
                return this.getItems(filterVal);
            case "Not Started":
                return this.getItems(filterVal);
            case "Completed":
                return this.getItems(filterVal);
            case "On Hold":
                return this.getItems(filterVal);
            default:
                return this.getItems(filterVal);
        }
    }

    public render(): React.ReactElement<ProjectsProps> {
        var notStarted = "Not Started";
        var inProg = "In Progress";
        var comp = "Completed";
        var onHold = "On Hold";
        var allPrj = "*";
        return <div>
            <div>
                <button onClick={() => this.progFilter(allPrj)}>
                    All
                </button>
                <button onClick={() => this.progFilter(inProg)}>
                    In Progress
                </button>
                <button onClick={() => this.progFilter(notStarted)}>
                    Not Started
                </button>
                <button onClick={() => this.progFilter(comp)}>
                    Completed
                </button>
                <button onClick={() => this.progFilter(onHold)}>
                    On Hold
                </button>
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