import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Stack, IStackTokens, FontSizes, IStackProps } from '@fluentui/react';
import { PrimaryButton } from '@fluentui/react/lib/Button';

const stackTokens: IStackTokens = { childrenGap: 20 };

export default class ConfirmationDialog extends BaseDialog {
    public paramFromDialog:string; 
    public message: string;

    public render() : void {    

        ReactDOM.render( 
            <div style={{ padding: 20 }}>
                <Stack tokens={stackTokens}>  
                    <div style={{ fontSize: FontSizes.size20 }} dangerouslySetInnerHTML={{__html: "Are you sure?"}}></div>
                    <div style={{ fontSize: FontSizes.size14, marginBottom: "30px" }} dangerouslySetInnerHTML={{__html: this.message}}></div>  
                </Stack>
                <Stack tokens={stackTokens} horizontal style={{justifyContent: 'flex-end'}}>                
                    <Stack.Item>
                        <PrimaryButton id="OkButton" text="Yes, please proceed" />  
                    </Stack.Item>
                    <Stack.Item>
                        <PrimaryButton id="CancelButton" text="No, cancel" />  
                    </Stack.Item>
                </Stack>
            </div>
            , 
            this.domElement);     

        this._setButtonEventHandlers();    
    }

    private _setButtonEventHandlers(): void {    
        const webPart: ConfirmationDialog = this;    
        this.domElement.querySelector('#OkButton').addEventListener('click', () => {   
             this.paramFromDialog = "OK";   
             this.close();  
        });  
        this.domElement.querySelector('#CancelButton').addEventListener('click', () => {   
            this.close();  
            // location.reload();
        });
    } 

    public async hideDialog() : Promise<void> {
        await this.close();
        return Promise.resolve();
    }
  
    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: true
        };
    }
  
    protected onAfterClose(): void {
        super.onAfterClose();
        
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
    
  }
