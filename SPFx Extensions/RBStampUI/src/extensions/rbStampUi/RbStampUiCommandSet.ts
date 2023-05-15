import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'RbStampUiCommandSetStrings';
import ConfirmationDialog from '../../components/ConfirmationDialog';
import { stampPDF } from './stampPDF';
import ProgressBarDialog from '../../components/ProgressBarDialog';
import { azureFuncProcess } from './azureFuncProcess';
import { getGraph, getSP } from '../pnpjsConfig';
import { SPFI } from '@pnp/sp';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IRbStampUiCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'RbStampUiCommandSet';

export default class RbStampUiCommandSet extends BaseListViewCommandSet<IRbStampUiCommandSetProperties> {

  private _isInProgress: boolean = false;
  private sp: SPFI;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized RbStampUiCommandSet');

    this.sp = getSP(this.context);
    getGraph(this.context);

    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const stampPDFCmd: Command = this.tryGetCommand('stampPDF');

    if (stampPDFCmd) {
      stampPDFCmd.visible = this.stampPDFIsActive(event);
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    if (this._isInProgress) { return; }
    this._isInProgress = true;

    let azureFunc: azureFuncProcess = null;


    switch (event.itemId) {

      case 'stampPDF':
        azureFunc = new stampPDF(this.context, event);

        const dialog: ConfirmationDialog = new ConfirmationDialog();
        var filename = event.selectedRows[0].getValueByName("FileRef");
        filename = filename.substring(filename.lastIndexOf("/") + 1);
        dialog.message = `The file ${filename} will be stamped.`;

        dialog.show().then(() => {

          if (dialog.paramFromDialog == "OK") {

            var id = event.selectedRows[0].getValueByName("ID");
            this.sp.web.lists.getById(this.context.pageContext.list.id.toString()).items.getById(id).select("hasStamp")<Array<any>>()
              .then((result: any) => {
                if (result["hasStamp"]) {
                  Dialog.alert("The selected pdf file has already a stamp.");
                  this._isInProgress = false;
                }
                else {
                  let progressBar: ProgressBarDialog = new ProgressBarDialog({ isBlocking: true });
                  progressBar.message = "Please wait. PDF Stamp is being generated.";
                  progressBar.show();

                  azureFunc.execute().then((result: string) => {
                    progressBar.hideDialog().then(() => {
                      if (result == "OK") //everything went well
                      {
                        Dialog.alert("The operation has been completed successfully.");
                        this._isInProgress = false;
                      }
                      else {
                        Dialog.alert(result);
                        this._isInProgress = false;
                      }

                    });

                  }).finally(() => {
                    this._isInProgress = false;
                  });
                }

              });

          }

        });
        this._isInProgress = false;
        break;
      default:
        throw new Error('Unknown command');
        this._isInProgress = false;
    }
  }

  private stampPDFIsActive(event: IListViewCommandSetListViewUpdatedParameters): boolean {
    if (event.selectedRows.length == 1) {
      // var id = event.selectedRows[0].getValueByName("ID"); 
      // console.log("The selected ID is: " + id);
      // var docIDUrl = event.selectedRows[0].getValueByName("_dlc_DocIdUrl"); 
      // console.log("The selected Document ID Url is: " + docIDUrl);
      // var docID = docIDUrl.substring(docIDUrl.lastIndexOf("=") + 1);
      // console.log("The selected Document ID is: " + docID);
      // var created = event.selectedRows[0].getValueByName("Created"); 
      // console.log("The selected Created Date is: " + created);
      // var dt = created.substring(0, created.indexOf(" "));
      // var year = dt.substring(dt.lastIndexOf("/") + 1);
      // console.log("The selected Year is: " + year);

      var filename = event.selectedRows[0].getValueByName("FileRef")
      console.log("The selected Filename is: " + filename);
      var extension = (filename.substring(filename.lastIndexOf(".") + 1)).toLowerCase();
      console.log("The selected Filename Extension is: " + extension);

      // sp.web.lists.getById(this.context.pageContext.list.id.toString()).items.getById(id).select("FileLeafRef").get<Array<any>>()
      // .then((result: any) => {
      //   if (result["FileLeafRef"]) {
      //      console.log("The selected Filename is: " + result["FileLeafRef"]);
      //      var extension = (result["FileLeafRef"].substring(result["FileLeafRef"].lastIndexOf(".") + 1)).toLowerCase();
      //      console.log("The selected Filename Extension is: " + extension);
      //   }
      // });

      //let permission = new SPPermission(this.context.pageContext.list.permissions.value);

      if (extension === "pdf") {
        return true;
      }
      else {
        return false;
      }
      //return extension == "pdf"; //&& permission.hasPermission(SPPermission.addListItems);
    }
    else {
      return false;
    }
  }
}
