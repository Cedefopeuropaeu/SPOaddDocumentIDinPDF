import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Stack, IStackTokens, FontSizes } from '@fluentui/react';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';

const stackTokens: IStackTokens = { childrenGap: 20 };

export default class ProgressBarDialog extends BaseDialog {
    public message: string;
     
    public render() : void {    
              
       ReactDOM.render( 
        <div style={{ padding: 50, minWidth: 500,  fontSize: FontSizes.size14 }}>
            <Stack tokens={stackTokens}>              
                <ProgressIndicator label={this.message} />
            </Stack>
        </div>
        , 
        this.domElement);   
        
        //slightly change the height for a better look and feel 
        var dialogElement = document.querySelector(".ms-Dialog-main") as HTMLElement;
        if (dialogElement) {
            dialogElement.style.height = dialogElement.style.minHeight = "150px";
        }    
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
