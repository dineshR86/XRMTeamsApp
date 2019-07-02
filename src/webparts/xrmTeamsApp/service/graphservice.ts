import { MSGraphClientFactory, MSGraphClient } from "@microsoft/sp-http";
import { Icaseitem } from "../model/Icaseitem";
import { Ilookupitem } from "../model/Ilookupitem";

export class graphservice{
    private _contextGraph:MSGraphClientFactory;
    private _siteid:string="cloudmission.sharepoint.com,fe171266-80d5-48e2-aac1-dd25051f3418,b1aa755f-a790-4cc3-9c20-a83dc3e92428";
    private _listid:string="3370e94e-b0a6-43da-8a86-64e7470ca1dc";
    private _categorylistid:string="8a9bd817-d9b5-467f-a8dc-6cc0161973f9";
    private _statuseslistid:string="b75dfe7d-5983-4e01-bf3f-b8cbdf39902a";
    private _clientslistid:string="b1a75ae4-3676-435a-929c-9080f6508510";
    
    //private _querystring:string="$expand=fields($select=Title,Column1,Column2,Column3,id)&$select=id,fields";

    constructor(contextGraph:MSGraphClientFactory){
        console.log("service constructor");
        this._contextGraph=contextGraph;
    }

    public GetListItems():Promise<any>{
        //debugger;
        let queryurl:string=`sites/${this._siteid}/lists/${this._listid}/items?select=id,fields/Title,fields/Deadline,fields/Responsible,fields/billable,fields/ClientLookupid,fields/StatusLookupid,fields/categorylookupid&expand=fields`;

        return this._contextGraph.getClient().then((client:MSGraphClient)=>{
            console.log("From client:", client);
            return client.api(queryurl).get().then((response)=>{
                console.log("From graph ",response);
                return response;
            });
        })
        .catch((error: any) => {
            console.log("Error: ", error);
        });
    }

    public GetClients():Promise<any>{
        //debugger;
        let queryurl:string=`sites/${this._siteid}/lists/${this._clientslistid}/items?select=id,fields/Title&expand=fields`;

        return this._contextGraph.getClient().then((client:MSGraphClient)=>{
            return client.api(queryurl).get().then((response)=>{
                let clientitems:Ilookupitem[]=[];
                response.value.forEach((item) => {
                    let nitem:Ilookupitem={
                        Id:item.id,
                        Title:item.fields.Title
                    };                   
                    clientitems.push(nitem);
                });
                return clientitems;
            });
        })
        .catch((error: any) => {
            console.log("Error: ", error);
        });
    }

    public GetCategory():Promise<any>{
        //debugger;
        let queryurl:string=`sites/${this._siteid}/lists/${this._categorylistid}/items?select=id,fields/Title&expand=fields`;

        return this._contextGraph.getClient().then((client:MSGraphClient)=>{
            return client.api(queryurl).get().then((response)=>{
                let categoryitems:Ilookupitem[]=[];
                response.value.forEach((item) => {
                    let nitem:Ilookupitem={
                        Id:item.id,
                        Title:item.fields.Title
                    };                   
                    categoryitems.push(nitem);
                });
                return categoryitems;
            });
        })
        .catch((error: any) => {
            console.log("Error: ", error);
        });
    }

    public GetStatuses():Promise<any>{
        //debugger;
        let queryurl:string=`sites/${this._siteid}/lists/${this._statuseslistid}/items?select=id,fields/Title&expand=fields`;

        return this._contextGraph.getClient().then((client:MSGraphClient)=>{
            return client.api(queryurl).get().then((response)=>{
                let statusesitems:Ilookupitem[]=[];
                response.value.forEach((item) => {
                    let nitem:Ilookupitem={
                        Id:item.id,
                        Title:item.fields.Title
                    };                   
                    statusesitems.push(nitem);
                });
                return statusesitems;
            });
        })
        .catch((error: any) => {
            console.log("Error: ", error);
        });
    }

    public PostXRMCases(xrmcase:any):Promise<any>{
        let queryurl:string=`https://graph.microsoft.com/v1.0/sites/${this._siteid}/lists/${this._listid}/items`;
        return this._contextGraph.getClient().then((client:MSGraphClient)=>{
            return client.api(queryurl).post(xrmcase).then((response)=>{
                return response;
            });
        })
        .catch((error: any) => {
            console.log("Error: ", error);
        });
    }
}