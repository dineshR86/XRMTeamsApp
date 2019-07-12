import * as React from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import * as moment from 'moment';
import { Row, Col, Layout, Table, Button, Input, Icon } from 'antd';
import Highlighter from 'react-highlight-words';
import './XrmTeamsApp.module.css';
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
  pagination: any;
  loading: boolean;
  _clients:Ilookupitem[];
  _statuses:Ilookupitem[];
  _category:Ilookupitem[];
  searchText:string;
}


export class XrmTeamsApp extends React.Component<IXrmTeamsAppProps, IXRMTeamAppState> {
  
  private searchInput;
  constructor(props: IXrmTeamsAppProps) {
    super(props);
    console.log("App constructor");
    this.state = {
      listItems: [],
      pagination: {},
      loading: false,
      _clients:[],
      _category:[],
      _statuses:[],
      searchText:''
    };
  }

  public componentDidMount() {
    this.setState({ loading: true });
    this.props.graphservice.GetXRMClients().then((resultc) => {
      this.setState({_clients:resultc});
      this.props.graphservice.GetXRMCategories().then((resultcg) => {
        this.setState({_category:resultcg});
        this.props.graphservice.GetXRMStatuses().then((results) => {
          this.setState({_statuses:results});
          this.props.graphservice.GetXRMCases({}).then((resultit) => {
            let items: Icaseitem[] = [];
            resultit.value.forEach((item) => {
              let clientT = this.state._clients.filter(i => i.Id == item.ClientId);
              let statusT = this.state._statuses.filter(i => i.Id == item.StatusId);
              let categoryT = this.state._category.filter(i => i.Id == item.CategoryId);
              const sdate: moment.Moment = moment(item.Deadline);
              let nitem: Icaseitem = {
                Id: item.Id,
                Title: item.Title,
                Deadline: sdate.format("MM/DD/YYYY"),
                //Responsible: item.fields.Responsible,
                billable: item.Billable,
                client: clientT.length > 0 ? clientT[0].Title : "",
                status: statusT.length > 0 ? statusT[0].Title : "",
                category: categoryT.length > 0 ? categoryT[0].Title : ""
              };

              items.push(nitem);
            });
            const pagination = { ...this.state.pagination };
            pagination.total = items.length;
            this.setState({
              listItems: items,
              loading: false,
              pagination
            });
          });
        });
      });
    });

  }

  public handleTableChange = (pagination, filters, sorter) => {
    debugger;
    const pager = { ...this.state.pagination };
    pager.current = pagination.current;
    this.setState({loading:true});
    this.props.graphservice.GetXRMCases(filters).then((resultit) => {
      let items: Icaseitem[] = [];
      resultit.value.forEach((item) => {
        let clientT = this.state._clients.filter(i => i.Id == item.ClientId);
        let statusT = this.state._statuses.filter(i => i.Id == item.StatusId);
        let categoryT = this.state._category.filter(i => i.Id == item.CategoryId);
        const sdate: moment.Moment = moment(item.Deadline);
        let nitem: Icaseitem = {
          Id: item.Id,
          Title: item.Title,
          Deadline: sdate.format("MM/DD/YYYY"),
          //Responsible: item.fields.Responsible,
          billable: item.Billable,
          client: clientT.length > 0 ? clientT[0].Title : "",
          status: statusT.length > 0 ? statusT[0].Title : "",
          category: categoryT.length > 0 ? categoryT[0].Title : ""
        };

        items.push(nitem);
      });
      this.setState({
        listItems: items,
        loading: false,
        pagination
      });
    });
  }

  public addnewcase=()=>{
    this.setState({loading:true});
    this.props.graphservice.GetXRMCases({}).then((resultit) => {
      let items: Icaseitem[] = [];
      resultit.value.forEach((item) => {
        let clientT = this.state._clients.filter(i => i.Id == item.ClientId);
        let statusT = this.state._statuses.filter(i => i.Id == item.StatusId);
        let categoryT = this.state._category.filter(i => i.Id == item.CategoryId);
        const sdate: moment.Moment = moment(item.Deadline);
        let nitem: Icaseitem = {
          Id: item.Id,
          Title: item.Title,
          Deadline: sdate.format("MM/DD/YYYY"),
          //Responsible: item.fields.Responsible,
          billable: item.Billable,
          client: clientT.length > 0 ? clientT[0].Title : "",
          status: statusT.length > 0 ? statusT[0].Title : "",
          category: categoryT.length > 0 ? categoryT[0].Title : ""
        };

        items.push(nitem);
      });
      const pagination = { ...this.state.pagination };
      pagination.total = items.length;
      this.setState({
        listItems: items,
        loading: false,
        pagination
      });
    });
  }

  public getColumnSearchProps = dataIndex => ({
    filterDropdown: ({ setSelectedKeys, selectedKeys, confirm, clearFilters }) => (
      <div style={{ padding: 8 }}>
        <Input
          ref={node => {
            this.searchInput = node;
          }}
          placeholder={`Search ${dataIndex}`}
          value={selectedKeys[0]}
          onChange={e => setSelectedKeys(e.target.value ? [e.target.value] : [])}
          onPressEnter={() => this.handleSearch(selectedKeys, confirm)}
          style={{ width: 188, marginBottom: 8, display: 'block' }}
        />
        <Button
          type="primary"
          onClick={() => this.handleSearch(selectedKeys, confirm)}
          icon="search"
          size="small"
          style={{ width: 90, marginRight: 8 }}
        >
          Search
        </Button>
        <Button onClick={() => this.handleReset(clearFilters)} size="small" style={{ width: 90 }}>
          Reset
        </Button>
      </div>
    ),
    filterIcon: filtered => (
      <Icon type="search" style={{ color: filtered ? '#1890ff' : undefined }} />
    ),
  });

  public handleSearch = (selectedKeys, confirm) => {
    debugger;
    confirm();
    this.setState({ searchText: selectedKeys[0] });
  };

  public handleReset = clearFilters => {
    clearFilters();
    this.setState({ searchText: '' });
  };

  public render(): React.ReactElement<IXrmTeamsAppProps> {
    const { Header, Content, Footer } = Layout;

    const columns = [
      {
        title: 'ID',
        dataIndex: 'Id',
        key: 'Id',
      },
      {
        title: 'Title',
        dataIndex: 'Title',
        key: 'Title',
        ...this.getColumnSearchProps('Title'),
      },
      {
        title: 'Client',
        dataIndex: 'client',
        key: 'client',
      },
      {
        title: 'Status',
        dataIndex: 'status',
        key: 'status',
        filters: [{ text: 'New', value: '1' }, { text: 'Closed', value: '5' }, { text: 'In Progress', value: '2' }, { text: 'Awaiting Internal Response', value: '3' },{ text: 'Awaiting Other', value: '4' }]
      },
      {
        title: 'Category',
        dataIndex: 'category',
        key: 'category',
      }
    
    ];
    
    

    return (
      <div>
        <Layout className="layout">
          <Header>
            <div className="logo" />
            <h4 className="xrmheader">XRM Teams App</h4>
          </Header>
          <Content style={{ padding: '0 50px' }}>
            <div className="gutter-example">
              <Row gutter={16}>
                <Col className="gutter-row" span={24}>
                  <div className="gutter-box">
                    <Xrmitemform clients={this.state._clients} status={this.state._statuses} category={this.state._category} graphservice={this.props.graphservice} addcase={this.addnewcase} />
                    <Table dataSource={this.state.listItems} columns={columns} pagination={this.state.pagination} loading={this.state.loading} onChange={this.handleTableChange} />
                  </div>
                </Col>
              </Row>
            </div>
          </Content>
          <Footer style={{ textAlign: 'center' }}>Teams Apps desgined by Thiru</Footer>
        </Layout>
      </div>
    );
  }
}
