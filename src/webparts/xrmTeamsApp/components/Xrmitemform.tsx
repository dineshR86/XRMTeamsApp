import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Ilookupitem } from '../model/Ilookupitem';
import { Icaseitem } from '../model/Icaseitem';
import { graphservice } from '../service/graphservice';

export interface XrmitemformProps {
  isNewForm: boolean;
  clients: Ilookupitem[];
  status: Ilookupitem[];
  category: Ilookupitem[];
  graphservice: graphservice;
}

export interface XrmitemformState {
  title: string;
  client: string;
  status: string;
  category: string;
  billable: boolean;
  isSuccess:boolean;
  isError:boolean;
}

export class Xrmitemform extends React.Component<XrmitemformProps, XrmitemformState>{

  constructor(props: XrmitemformProps) {
    super(props);
    console.log("App constructor");
    this.state = {
      title: "",
      client: "",
      status: "",
      category: "",
      billable: false,
      isSuccess:false,
      isError:false
    };

  }

  @autobind
  public handleTitleChange(event) {
    this.setState({ title: event.target.value });
  }

  @autobind
  public handleClientChange(event) {
    this.setState({ client: event.target.value });
  }

  @autobind
  public handleStatusChange(event) {
    this.setState({ status: event.target.value });
  }

  @autobind
  public handleCategoryChange(event) {
    this.setState({ category: event.target.value });
  }

  @autobind
  public handleBillableChange(event) {
    this.setState({ billable: event.target.checked });
  }

  @autobind
  public handleSubmit(event) {
    debugger;
    const { title, client, status, category, billable } = this.state;
    const xrmcase = {
      fields: {
        Title: title,
        Billable: billable,
        ClientLookupId: client,
        StatusLookupId: status,
        CategoryLookupId: category
      }
    };

    this.props.graphservice.PostXRMCases(xrmcase).then((result) => {
      console.log("post success: ", result);
      this.setState({
        title: "",
        billable: false,
        client: "",
        status: "",
        category: "",
        isSuccess:true
      });
    }).catch((error: any) => {
      console.log("Error: ", error);
    });

    //event.preventDefault();
  }

  public render(): React.ReactElement<XrmitemformProps> {
    return (
      <div>
        {this.state.isSuccess ?<div className="alert alert-success" role="alert">Case created successfully! refresh the app</div>:""}
        {this.state.isError?<div className="alert alert-danger" role="alert">A simple danger alert—check it out!</div>:""}
        <div className="form-group">
          <label>Title</label>
          <input type="text" className="form-control" id="exampleFormControlInput1" placeholder="Enter the Case Title" value={this.state.title} onChange={this.handleTitleChange} />
        </div>
        <div className="form-group">
          <label>Client</label>
          <select className="form-control" id="exampleFormControlSelect1" value={this.state.client} onChange={this.handleClientChange} >
            <option value="" selected>-select-</option>
            {this.props.clients.map((item, index) => <option value={item.Id} key={index}>{item.Title}</option>)}
          </select>
        </div>
        <div className="form-group">
          <label>Status</label>
          <select className="form-control" id="exampleFormControlSelect2" value={this.state.status} onChange={this.handleStatusChange}>
            <option value="" selected>-select-</option>
            {this.props.status.map((item, index) => <option value={item.Id} key={index}>{item.Title}</option>)}
          </select>
        </div>
        <div className="form-group">
          <label>Category</label>
          <select className="form-control" id="exampleFormControlSelect2" value={this.state.category} onChange={this.handleCategoryChange}>
            <option value="" selected>-select-</option>
            {this.props.category.map((item, index) => <option value={item.Id} key={index}>{item.Title}</option>)}
          </select>
        </div>
        <div className="custom-control custom-switch">
          <input type="checkbox" className="custom-control-input" id="customSwitch1" checked={this.state.billable} onChange={this.handleBillableChange} />
          <label className="custom-control-label" htmlFor="customSwitch1">Billable</label>
        </div>
        <button type="button" className="btn btn-primary" onClick={this.handleSubmit}>Submit</button>
      </div>
    );
  }
}