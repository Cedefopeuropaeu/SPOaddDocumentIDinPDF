import { IConfigurationParameter } from '../entities/IConfigurationParameter';
import { virtual } from '@microsoft/decorators';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { MSGraphClient } from '@microsoft/sp-http'; 
import { IListViewCommandSetExecuteEventParameters, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import { GraphFI } from '@pnp/graph';
import { SPFI } from '@pnp/sp';
import { getGraph, getSP } from '../pnpjsConfig';



export interface IFileInfo {
    FilePath: string;
    FileName: string;
}

export abstract class azureFuncProcess
{   
    protected _azureFunctionURL : IConfigurationParameter;
    protected _clientAPI : IConfigurationParameter;
    protected _context: ListViewCommandSetContext;   
    protected _event: IListViewCommandSetExecuteEventParameters;
    private graph: GraphFI;
    private sp: SPFI;

    constructor(context: any, event: IListViewCommandSetExecuteEventParameters){  
        this._context = context;  
        this._event = event;     
        this.graph = getGraph();
        this.sp = getSP();
        let timeStamp = new Date().toLocaleDateString("en-UK", { hour: '2-digit', minute: '2-digit' }).replace(",","");     
    }

    protected abstract get azureFuncName() : string;

    @virtual
    public async azureFuncExecutionStep() : Promise<string> {       
        
        //let updateBatch = sp.web.createBatch();

        var id = this._event.selectedRows[0].getValueByName("ID"); 
        var libUrl = this._context.pageContext.list.serverRelativeUrl;
        var docIDUrl =  this._event.selectedRows[0].getValueByName("_dlc_DocIdUrl"); 
        var fileRef: string = this._event.selectedRows[0].getValueByName("FileRef"); 
        var docID = docIDUrl.substring(docIDUrl.lastIndexOf("=") + 1);
        var FileLeafRef =  this._event.selectedRows[0].getValueByName("FileLeafRef");
        var tenantid: string = this._context.pageContext.aadInfo.tenantId['_guid'];
        var siteUrl = this._context.pageContext.site.absoluteUrl;
        var extension = (FileLeafRef.substring(FileLeafRef.lastIndexOf(".") + 1)).toLowerCase();
        const body: string = JSON.stringify({
            'FileLeafRef' : FileLeafRef,
            'fileRef' : fileRef,
            'tenantId': tenantid,
            'siteUrl': siteUrl,
            'docID': docID,
            'libUrl': libUrl,
            'id': id
        });
        
        console.log(this._event.selectedRows[0]);
        console.log(body);
        const options: IHttpClientOptions = { body: body };
        //const client = await this._context.aadHttpClientFactory.getClient('api://05e03e4f-f753-4c38-a27d-fbfe9816826d');   // MSDNCOLLAB
        const client = await this._context.aadHttpClientFactory.getClient(this._clientAPI.Text);     //CEDEFOPDEV
        
        //let url:string = this._azureFunctionURL.Text + `?siteUrl=${siteUrl}&id=${id}&filename=${filename}&docIDUrl=${docIDUrl}&filename=${filename}&extension=${extension}`;
        var response: HttpClientResponse = await client.post(this._azureFunctionURL.Text, AadHttpClient.configurations.v1, options);
        if (response.status === 200)
        {
            return Promise.resolve('OK');
        }
        else
        {
            var returnvalue: string  = 'Error code : ' + response.status.toString() + ' '  + await response.text();
            return Promise.resolve(returnvalue);
        }
    }

    public async execute() : Promise<string> {
        await this.initFunc();              
        return  await this.azureFuncExecutionStep();
    }

    protected async initFunc() : Promise<void> 
    {           
        this._azureFunctionURL = await this.getAzureFunctionURL();     
        this._clientAPI = await this.getClientAPI();   
        return Promise.resolve();
    }  

    public async getAzureFunctionURL() : Promise<IConfigurationParameter> 
    {
        const result = await this.sp.web.lists.getByTitle('Configuration Parameters')
            .items.getById(1)
            .select("Title", "Text")() as IConfigurationParameter; 
        
        return result;
    }
    public async getClientAPI() : Promise<IConfigurationParameter> 
    {
        const result = await this.sp.web.lists.getByTitle('Configuration Parameters')
            .items.getById(2)
            .select("Title", "Text")() as IConfigurationParameter; 
        
        return result;
    }

}