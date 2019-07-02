import * as React from 'react';
import { Icaseitem } from '../model/Icaseitem';

export interface XrmListitemProps{
    items:Icaseitem[];
}

export class XrmListitem extends React.Component<XrmListitemProps,{}>{

    public render():React.ReactElement<XrmListitemProps>{
        return(
            <table className="table table-striped table-bordered">
              <thead>
                <tr>
                  <th scope="col">#</th>
                  <th scope="col">Title</th>
                  <th scope="col">Client</th>
                  <th scope="col">Status</th>
                  <th scope="col">Category</th>
                  <th scope="col">Deadline</th>
                  <th scope="col">Billable</th>
                </tr>
              </thead>
              <tbody>
              {/* <td>{item.Responsible?item.Responsible[0].LookupValue:""}</td> */}
               {this.props.items.map((item,index)=>{
                   return <tr key={index}><th scope="row">{item.Id}</th><td>{item.Title}</td><td>{item.client}</td><td>{item.status}</td><td>{item.category}</td><td>{item.Deadline}</td><td>{item.billable?"Yes":"No"}</td></tr>;
               })}
              </tbody>
            </table>
        );
    }
}