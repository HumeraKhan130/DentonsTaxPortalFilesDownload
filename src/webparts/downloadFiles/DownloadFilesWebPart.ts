import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DownloadFilesWebPartStrings';
import DownloadFiles from './components/DownloadFiles';
import { IDownloadFilesProps } from './components/IDownloadFilesProps';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { IODataList } from '@microsoft/sp-odata-types';



export default class DownloadFilesWebPart extends BaseClientSideWebPart<IDownloadFilesProps> {


  public dropdownOptions: IPropertyPaneDropdownOption[];
  public listsFetched: boolean;
  
  protected get disableReactivePropertyChanges(): boolean {
    return true;
    }

  private fetchLists(url: string) : Promise<any> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
        return null;
      }
    });
}

private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
  var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`;

  return this.fetchLists(url).then((response) => {
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      response.value.map((list: IODataList) => {
          options.push( { key: list.Id, text: list.Title });
        console.log(list.Id +"  "+ list.Title);
      });
      this.context.propertyPane.refresh();
      return options;
  });
}

  public render(): void {

    
    const element: React.ReactElement<IDownloadFilesProps> = React.createElement(
      DownloadFiles,
      {
        context: this.context,
        siteUrl:this.context.pageContext.site.absoluteUrl,
        ImageUrl:this.properties.ImageUrl,
        PartnerList: "Partners",
        PartnerTaxFilesList:this.properties.PartnerTaxFilesList
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

   
    if (!this.listsFetched) {
      this.fetchOptions().then((response) => {
        this.dropdownOptions = response;
        this.listsFetched = true;
          
        // now refresh the property pane, now that the promise has been resolved..
        this.render();
        this.onDispose();
      });
   }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('ImageUrl', {
                  label: "Image URL",
                  multiline:false,
                  resizable:false,
                  value:"https://bdocollab.sharepoint.com/sites/SPFX_Test_humeraK/siteassets/notify.jpeg",
                  placeholder:"Please select the image"

                }),
                PropertyPaneDropdown('PartnersList', {
                  label: "Partners List",
                  options:this.dropdownOptions,
                  selectedKey:'91c779be-5cbc-4efe-baf4-9e1de5328487'
                }),
                PropertyPaneDropdown('PartnersTaxFilesList', {
                  label: "Partners Tax Files",
                  options:this.dropdownOptions

                })
              ]
            }
          ]
        }
      ]
    };
  }
}
