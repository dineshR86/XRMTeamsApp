import * as React from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import * as moment from 'moment';
import { graphservice } from '../service/graphservice';
import { Icaseitem } from '../model/Icaseitem';
import { XrmListitem } from './XrmListItem';
import { Xrmitemform } from './Xrmitemform';
import { Ilookupitem } from '../model/Ilookupitem';


export interface IXrmTeamsAppProps {
  description: string;
  teamsContext: microsoftTeams.Context;
  graphservice: graphservice;
}

export interface IXRMTeamAppState {
  listItems: Icaseitem[];
}

export class XrmTeamsApp extends React.Component<IXrmTeamsAppProps, IXRMTeamAppState> {
  private _clients: Ilookupitem[];
  private _statuses: Ilookupitem[];
  private _category: Ilookupitem[];

  constructor(props: IXrmTeamsAppProps) {
    super(props);
    console.log("App constructor");
    this.state = {
      listItems: []
    };
  }

  public componentDidMount() {
    console.log("componentdidmount");
    this.props.graphservice.GetClients().then((resultc) => {
      this._clients = resultc;
      console.log("Clients");
        this.props.graphservice.GetCategory().then((resultcg) => {
          this._category = resultcg;
          console.log("Category");
            this.props.graphservice.GetStatuses().then((results) => {
              this._statuses = results;
              console.log("Statuses");
                this.props.graphservice.GetListItems().then((resultit) => {
                  console.log("resul from component ", resultit);
                  let items: Icaseitem[] = [];
                  resultit.value.forEach((item) => {
                    debugger;
                    let clientT=this._clients.filter(i=>i.Id==item.fields.ClientLookupId);
                    let statusT=this._statuses.filter(i=>i.Id==item.fields.StatusLookupId);
                    let categoryT=this._category.filter(i=>i.Id==item.fields.CategoryLookupId);
                    const sdate:moment.Moment=moment(item.fields.Deadline);
                    let nitem: Icaseitem = {
                      Id: item.id,
                      Title: item.fields.Title,
                      Deadline: sdate.format("MM/DD/YYYY"),
                      Responsible: item.fields.Responsible,
                      billable: item.fields.Billable,
                      client: clientT.length>0?clientT[0].Title:"",
                      status: statusT.length>0?statusT[0].Title:"",
                      category: categoryT.length>0?categoryT[0].Title:""
                    };
            
                    items.push(nitem);
                  });
                  console.log("items to render: ",items);
                  this.setState({
                    listItems: items
                  });
                });
            });
        });
    });
    
  }

  public render(): React.ReactElement<IXrmTeamsAppProps> {
    //debugger;
    console.log("render");
    if (this._clients) {
      console.log("clients from render: ", this._clients);
    }

    const { listItems } = this.state;
    if (listItems.length <= 0) {
      return <div>Fetching Data</div>;
    }

    return (
      <div className="container-fluid border" >
        <div className="rows">
          <div className="col" style={{ textAlign: "center" }}>
            <p className="h4">XRM Teams App</p>
          </div>
        </div>
        <div className="row">
          <div className="col-8" style={{overflow:'hidden'}}>
            <XrmListitem items={this.state.listItems} />
          </div>
          <div className="col-4">
            <Xrmitemform isNewForm={false} clients={this._clients} status={this._statuses} category={this._category} graphservice={this.props.graphservice} />
          </div>
        </div>
      </div>
    );
  }
}
