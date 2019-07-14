import * as React from 'react';
import { Modal,Select,DatePicker,Radio,Icon,Button } from 'antd';
import { Ilookupitem } from '../model/Ilookupitem';
import { Icaseitem } from '../model/Icaseitem';
import { graphservice } from '../service/graphservice';

const { Option } = Select;

export interface XrmitemformProps {
  isNewForm?: boolean;
  clients?: Ilookupitem[];
  status?: Ilookupitem[];
  category?: Ilookupitem[];
  graphservice?: graphservice;
  addcase?:any;
}

export interface XrmitemformState {
  modalvisible: boolean;
  modalsave: boolean;
  title?:string;
  clientid?:string;
  statusid?:string;
  categoryid?:string;
  deadline?:string;
  billable?:boolean;
  
}

export class Xrmitemform extends React.Component<XrmitemformProps, XrmitemformState>{

  constructor(props: XrmitemformProps) {
    super(props);
    this.state = {
      modalvisible: false,
      modalsave: false,
      billable:false
    };

  }

  public handleOk = (e) => {
    debugger;
    this.setState({ modalsave: true });
    const { title, clientid, statusid, categoryid, billable } = this.state;
    const xrmcase = {
        Title: title,
        Billable: billable,
        ClientId: clientid,
        StatusId: statusid,
        CategoryId: categoryid
    };

    this.props.graphservice.AddXRMCase(xrmcase).then((result) => {
      console.log("post success: ", result);
      this.props.addcase();
      this.setState({
        title: "",
        billable: false,
        clientid: "",
        statusid: "",
        categoryid: "",
        modalsave:false,
        modalvisible:false
      });
    }).catch((error: any) => {
      console.log("Error: ", error);
    });
  }

  public handleCancel = (e) => {
    console.log(e);
    this.setState({
      modalvisible: false
    });
  }

  public showModal = () => {
    this.setState({
      modalvisible: true
    });
  }

  public titleChange=(e)=>{
    //debugger;
    //console.log(e.currentTarget.value);
    this.setState({title:e.currentTarget.value});

  }

  public caseChange=(value)=>{
    console.log("Case selected: ",value);
    this.setState({clientid:value});
  }

  public statusChange=(value)=>{
    //console.log("Case selected: ",value);
    this.setState({statusid:value});
  }

  public categoryChange=(value)=>{
    //console.log("Case selected: ",value);
    this.setState({categoryid:value});
  }

  public onDateChange=(date,dateString)=>{
    console.log(date, dateString);
    this.setState({deadline:date});
  }

  public onBillableChange=(e)=>{
    this.setState({billable:e.target.value});
  }

  

  public render(): React.ReactElement<XrmitemformProps> {
    return (
      <div>
        <div>
        {/* <Icon type="file-add" theme="twoTone"  style={{fontSize:'30px',float:'right'}} /> <Icon type="plus-square" theme="twoTone" /> */}
        <Button icon="plus-square" onClick={this.showModal} style={{marginLeft:"90%"}} >Add Case</Button>
        </div>
        <Modal
          title="New Case"
          visible={this.state.modalvisible}
          onOk={this.handleOk}
          onCancel={this.handleCancel}
          okText="Submit"
          confirmLoading={this.state.modalsave}
        >
          <div className="ant-form ant-form-vertical">
            <div className="ant-row ant-form-item">
              <div className="ant-col ant-form-item-label">
                <label>Title</label>
              </div>
              <div className="ant-col ant-form-item-control-wrapper">
                <div className="ant-form-item-control">
                  <span className="ant-form-item-children">
                    <input type="text" className="ant-input" placeholder="Case Title" onChange={this.titleChange} />
                  </span>
                </div>
              </div>
            </div>
            <div className="ant-row ant-form-item">
              <div className="ant-col ant-form-item-label">
                <label>Client</label>
              </div>
              <div className="ant-col ant-form-item-control-wrapper">
                <div className="ant-form-item-control">
                  <span className="ant-form-item-children">
                   <Select placeholder="Please select a client" style={{ width: 120 }} onChange={this.caseChange}>
                     {this.props.clients.map((client:Ilookupitem,index)=> <Option value={client.Id} key={index}>{client.Title}</Option>)}
                   </Select>
                  </span>
                </div>
              </div>
            </div>
            <div className="ant-row ant-form-item">
              <div className="ant-col ant-form-item-label">
                <label>Status</label>
              </div>
              <div className="ant-col ant-form-item-control-wrapper">
                <div className="ant-form-item-control">
                  <span className="ant-form-item-children">
                   <Select placeholder="Please select a status" style={{ width: 120 }} onChange={this.statusChange}>
                   {this.props.status.map((sta:Ilookupitem,index)=> <Option value={sta.Id} key={index}>{sta.Title}</Option>)}
                   </Select>
                  </span>
                </div>
              </div>
            </div>
            <div className="ant-row ant-form-item">
              <div className="ant-col ant-form-item-label">
                <label>Categories</label>
              </div>
              <div className="ant-col ant-form-item-control-wrapper">
                <div className="ant-form-item-control">
                  <span className="ant-form-item-children">
                   <Select placeholder="Please select a Categorie" style={{ width: 120 }} onChange={this.categoryChange}>
                    {this.props.category.map((catg:Ilookupitem,index)=> <Option value={catg.Id} key={index}>{catg.Title}</Option>)}
                   </Select>
                  </span>
                </div>
              </div>
            </div>
            <div className="ant-row ant-form-item">
              <div className="ant-col ant-form-item-label">
                <label>Deadline</label>
              </div>
              <div className="ant-col ant-form-item-control-wrapper">
                <div className="ant-form-item-control">
                  <span className="ant-form-item-children">
                   <DatePicker onChange={this.onDateChange} />
                  </span>
                </div>
              </div>
            </div>
            <div className="ant-row ant-form-item">
              <div className="ant-col ant-form-item-label">
                <label>Billable</label>
              </div>
              <div className="ant-col ant-form-item-control-wrapper">
                <div className="ant-form-item-control">
                  <span className="ant-form-item-children">
                   <Radio.Group value={this.state.billable} onChange={this.onBillableChange}>
                     <Radio value={false}>No</Radio>
                     <Radio value={true}>Yes</Radio>
                     </Radio.Group>
                  </span>
                </div>
              </div>
            </div>
          </div>
        </Modal>
      </div>
    );
  }
}