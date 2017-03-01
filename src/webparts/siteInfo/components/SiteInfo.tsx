import * as React from 'react';
import * as fabric from 'office-ui-fabric-react'; // should just import needed modules for production use
import * as pnp from 'sp-pnp-js';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { ISiteInfoWebPartProps } from '../ISiteInfoWebPartProps';

export interface ISiteInfoProps extends ISiteInfoWebPartProps {
}

export default class SiteInfo extends React.Component<ISiteInfoProps, {}> {
  constructor(props: ISiteInfoProps) {
    super(props);
    // set initial state
    this.state = { Title: "", SpinnerClass: "", SpinnerType: fabric.SpinnerType.large, SpinnerVisible: true, PanelVisible: false, ButtonClass: "ms-Button--primary", Lists: [] };
  }

  public componentDidMount(): void {
    var that: any = this; // save 'this' so it is available from within the closure
    if (!(Environment.type === EnvironmentType.Local)) {
      // make sure we are on the correct web...
      var web = new pnp.Web(this.props.self.context.pageContext.web.absoluteUrl);
      web.get().then(result => {
        that.setState({ Title: result.Title });
      });
      web.lists.get().then(result => {
        var lists: any = result.map(r => r.Title);
        that.setState({ Lists: lists });
      });
    }
    else { // running locally - use test data
      that.setState({
        Title: "My Site Title",
        Lists: ["Documents", "My List 1", "My List 2"],
      });
    }
  }

  private renderUser(): JSX.Element {
    if (!this.props.showUser) return null;
    var user: string = this.props.self.context.pageContext.user.displayName;
    var login: string = this.props.self.context.pageContext.user.loginName;
    return (<h3>User: {user} ({login}) </h3>);
  }

  private renderItem(item: string): JSX.Element {
    return <h3><i className="ms-Icon ms-Icon--star" aria-hidden="true"></i> { item }</h3>;
  }

  private renderLists(): JSX.Element {
    if (this.props.showLists) return (
      <div>
        <h2>The site contains the following lists: </h2>
        <fabric.List items={this.state["Lists"]} onRenderCell={ this.renderItem } />
      </div>);
    return undefined;
  }

  private despinner_click(): void {
    this.setState({ SpinnerVisible: false, ButtonClass: "ms-Button is-disabled" });
  }

  public render(): JSX.Element {
    return (
      <div>
        <h1 className="ms-bgColor-themeLighter">{this.props.description + this.state["Title"]} </h1>
        <fabric.Button onClick={() => this.setState({ PanelVisible: true })} className="ms-Button"><i className="ms-Icon ms-Icon--listCheckbox ms-fontSize-xxl" aria-hidden="true"></i> More...</fabric.Button>
        <fabric.Panel isOpen={this.state["PanelVisible"]} onDismiss={() => this.setState({ PanelVisible: false }) }>
          <h2>More information for {this.props.self.context.pageContext.web.title}</h2>
          <ul>
            <li>serverRelativeUrl = {this.props.self.context.pageContext.web.serverRelativeUrl}</li>
            <li>absoluteUrl = {this.props.self.context.pageContext.web.absoluteUrl}</li>
          </ul>
        </fabric.Panel>
        {this.renderUser() }
        {this.renderLists() }
        {this.state["SpinnerVisible"] ? <fabric.Spinner type={this.state["SpinnerType"]} /> : <span/> }
        <hr />
        <fabric.Button onClick={() => this.despinner_click() } className={this.state["ButtonClass"]}>Dismiss Gratuitous Spinner</fabric.Button>
        {Environment.type === EnvironmentType.Local ? <h3 className="ms-bgColor-error">Running locally with mock data</h3> : <span/> }
      </div>
    );
  }
}