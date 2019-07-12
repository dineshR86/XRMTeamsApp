import { MSGraphClientFactory, MSGraphClient,SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from "@microsoft/sp-http";
import { Icaseitem } from "../model/Icaseitem";
import { Ilookupitem } from "../model/Ilookupitem";

export class graphservice{
    private _contextGraph:MSGraphClientFactory;
    private _spclient:SPHttpClient;
    private _weburl:string="https://cloudmission.sharepoint.com/sites/xrmtrial/";
    private _siteid:string="cloudmission.sharepoint.com,fe171266-80d5-48e2-aac1-dd25051f3418,b1aa755f-a790-4cc3-9c20-a83dc3e92428";
    private _listid:string="3370e94e-b0a6-43da-8a86-64e7470ca1dc";
    private _categorylistid:string="8a9bd817-d9b5-467f-a8dc-6cc0161973f9";
    private _statuseslistid:string="b75dfe7d-5983-4e01-bf3f-b8cbdf39902a";
    private _clientslistid:string="b1a75ae4-3676-435a-929c-9080f6508510";
    
    //private _querystring:string="$expand=fields($select=Title,Column1,Column2,Column3,id)&$select=id,fields";

    constructor(contextGraph:MSGraphClientFactory,spclient:SPHttpClient){
        console.log("service constructor");
        this._contextGraph=contextGraph;
        this._spclient=spclient;
    }

    public GetXRMCases(filters?:any):Promise<any>{
        debugger;
        let casefilter:string="";
        if(typeof filters.status != "undefined"){
        for(let i in filters.status){
            if(i=="0"){
                casefilter=`$filter=StatusId eq ${filters.status[i]}`;
            }else{
            casefilter= casefilter.concat(`or StatusId eq ${filters.status[i]}`)
        }
        }
    }

    if(typeof filters.Title != "undefined"){
        for(let i in filters.Title){
            if(i=="0"){
                casefilter=`$filter=substringof('${filters.Title[0]}',Title)`;
            }
        }
    }
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._listid}')/items?$select=Title,Id,ClientId,StatusId,CategoryId,Deadline,Billable&$orderby=Id desc&${casefilter}`;
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
               return data;
            }).catch((ex) => {
                console.log("Error while fetching XRMCases: ", ex);
                throw ex;
            });
    }

    public GetXRMClients():Promise<Ilookupitem[]>{
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._clientslistid}')/items?$select=Title,Id`;
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let clientitems:Ilookupitem[]=[];
                data.value.forEach((item) => {
                    let nitem:Ilookupitem={
                        Id:item.ID,
                        Title:item.Title
                    };                   
                    clientitems.push(nitem);
                });
                return clientitems;
            }).catch((ex) => {
                console.log("Error while fetching clients: ", ex);
                throw ex;
            });
    }

    public GetXRMStatuses():Promise<Ilookupitem[]>{
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._statuseslistid}')/items?$select=Title,Id`;
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let statusesitems:Ilookupitem[]=[];
                data.value.forEach((item) => {
                    let nitem:Ilookupitem={
                        Id:item.ID,
                        Title:item.Title
                    };                   
                    statusesitems.push(nitem);
                });
                return statusesitems;
            }).catch((ex) => {
                console.log("Error while fetching Status: ", ex);
                throw ex;
            });
    }

    public GetXRMCategories():Promise<any>{
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._categorylistid}')/items?$select=Title,Id`;
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let categoryitems:Ilookupitem[]=[];
                data.value.forEach((item) => {
                    let nitem:Ilookupitem={
                        Id:item.ID,
                        Title:item.Title
                    };                   
                    categoryitems.push(nitem);
                });
                return categoryitems;
            }).catch((ex) => {
                console.log("Error while fetching clients: ", ex);
                throw ex;
            });
    }

    public AddXRMCase(xrmcase:any):Promise<any>{
        const addcaseurl:string=`${this._weburl}_api/web/lists(guid'${this._listid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            body:JSON.stringify(xrmcase)
        };

        return this._spclient.post(addcaseurl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });

    }

    // public PostXRMCases(xrmcase:any):Promise<any>{
    //     let queryurl:string=`https://graph.microsoft.com/v1.0/sites/${this._siteid}/lists/${this._listid}/items`;
    //     return this._contextGraph.getClient().then((client:MSGraphClient)=>{
    //         return client.api(queryurl).post(xrmcase).then((response)=>{
    //             return response;
    //         });
    //     })
    //     .catch((error: any) => {
    //         console.log("Error: ", error);
    //     });
    // }
}