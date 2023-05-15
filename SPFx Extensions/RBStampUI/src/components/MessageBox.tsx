import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Stack, IStackTokens, FontSizes } from '@fluentui/react';
import { PrimaryButton } from '@fluentui/react/lib/Button';

const stackTokens: IStackTokens = { childrenGap: 20 };

export default class MessageBox extends BaseDialog {
    public message: string;
     
    public render() : void {    
              
       ReactDOM.render( 
        <div style={{ padding: 50 }}>
            <Stack tokens={stackTokens}>
                <div style={{ fontSize: FontSizes.size14 }} dangerouslySetInnerHTML={{__html: this.message}}></div>      
                <Stack.Item align="end">
                    <PrimaryButton id="OkButton" text="OK" />  
                </Stack.Item>
            </Stack>
        </div>
        , 
        this.domElement);     
        this._setButtonEventHandlers();     
    }

    private _setButtonEventHandlers(): void {    
        const webPart: MessageBox = this;    
        this.domElement.querySelector('#OkButton').addEventListener('click', () => {    
                 
            this.close();  
         });   
    }   
  
    public getConfig(): IDialogConfiguration {
        return {
        isBlocking: false
        };
    }
  
    protected onAfterClose(): void {
        super.onAfterClose();
        
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
    
  }
