import { azureFuncProcess } from './azureFuncProcess';
import { IListViewCommandSetExecuteEventParameters } from '@microsoft/sp-listview-extensibility';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export class stampPDF extends azureFuncProcess
{
    constructor(context: ListViewCommandSetContext, event: IListViewCommandSetExecuteEventParameters){
       super(context, event);       
    }    

    protected get azureFuncName() : string {
        return "Stamp PDF";
    }

  
}   