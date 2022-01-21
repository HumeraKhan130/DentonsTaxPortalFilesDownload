import * as React from 'react';
import styles from './DownloadFiles.module.scss';
import { IDownloadFilesProps } from './IDownloadFilesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import * as jquery from 'jquery';
import { PropertyPaneSlider } from '@microsoft/sp-property-pane';


var currUser:string;
export default class DownloadFiles extends React.Component<IDownloadFilesProps, {}> {
 
  constructor(props: IDownloadFilesProps, state: {}){
    super(props);
    this.state = ({
      showDialog: true,
     
    });

    
  }

 

  public render(): React.ReactElement<IDownloadFilesProps> {

   console.log( this.props.PartnerList);
    return (
      <div className={ styles.downloadFiles }>
      <div className={ styles.container }>
        <div className={ styles.row }>
           {this.props.PartnerList}
              <a  id="DownloadFiles" onClick={ ()=> this._downloadFiles() }> 
                <img className={styles.image}  src={this.props.ImageUrl == undefined ? this.props.context.pageContext.site.absoluteUrl+"/siteassets/notify.jpeg" : this.props.ImageUrl}></img>
                <span className= {styles.span}>Download</span>
              </a>
        </div>
        <div className={ styles.row }>
                <div id='statusDiv' className={styles.statusDiv} hidden >
                          <p id='inProcess' hidden >Sending </p>
                          <p id='completeProcess' hidden>Email has been sent to all partners.</p>
                </div>
        </div>

      </div>
     
         
  </div> 
    );
  }
  

  private  _downloadFiles(): void { 

  
    /*****************************************************Get Partner name and Partner Tax Files*********************************************** */
     this.getPartnerWithAsyncAwait();
   

 
      
    
    /*
    */
  
  }


  public   getPartnerWithAsyncAwait() {
    /* console.log(this.props);
     const response = await this.props.context.spHttpClient.get(this.props.siteUrl + 
                     "/_api/web/lists/GetByTitle('" + this.props.PartnerList+"')/Items?$select=Partner/Title&$filter=Partner/Title eq '"+ this.props.context.pageContext.user.displayName+ "'&$expand=Partner",SPHttpClient.configurations.v1)
     const user: any = await  response.json();
     return user.value[0].Partner["Title"];*/
  var siteUrl = this.props.siteUrl;
   jquery.ajax({
       url: `${this.props.siteUrl}/_api/web/lists/GetByTitle('` + this.props.PartnerList+`')/Items?$select=Partner/Title,LinkTitle&$filter=Partner/Title eq '`+ this.props.context.pageContext.user.displayName+ `'&$expand=Partner`, 
       type: "GET",
       headers: { "Accept": "application/json; odata=verbose" ,"Content-Type": "application/json; " },
       success: function(items) {          
         getPartnerTaxFilesWithAsyncAwait(siteUrl, items.d.results[0].LinkTitle);
       }
     });
  
  
   }


   

}


function  getPartnerTaxFilesWithAsyncAwait(siteUrl, PartnerName) {
    
  var partnerfolder = PartnerName+" Tax Files"; 
  var currentYear= "";
 jquery.ajax({
   url: siteUrl + `/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27/sites/DentonsPartnerPortal-PMesiha/Partner Tax Files%27&RootFolder=/sites/DentonsPartnerPortal-PMesiha/Partner Tax Files/`+ partnerfolder , 
   type: "POST",
   contentType:"application/json;charset=utf-8",
   data: "{\"parameters\" : {\"RenderOptions\":4103}}",
   success: function(data,textStatus,jqXHR) {          
      
     
     var result = FormatArrayWithResult(jqXHR.responseJSON.ListData["Row"].sort(sortDesc)[0],jqXHR.responseJSON.ListSchema[".driveAccessToken"]);
      //console.log(result);
      currentYear = jqXHR.responseJSON.ListData["Row"][0].FileLeafRef;
      var xhr = new XMLHttpRequest();
      xhr.open("POST", "https://canadaeast1-mediap.svc.ms/transform/zip?cs=fFNQTw", true);
      xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded;");
      xhr.responseType = "blob";
    
      xhr.onreadystatechange = function() {
        if (xhr.readyState == XMLHttpRequest.DONE) {
         //   console.log(xhr.response);
            var blob = new Blob([xhr.response], { type: "application/zip" });
            //var blob = new Blob([zip]);
            var link = document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = PartnerName +"-"+ currentYear +".zip";
            link.click();
        }
      }

     //xhr.send("provider=spo&files=%7B%22items%22%3A%5B%7B%22name%22%3A%222021%22%2C%22size%22%3A%2219452%22%2C%22docId%22%3A%22https%3A%2F%2Fbdocollab.sharepoint.com%3A443%2F_api%2Fv2.0%2Fdrives%2Fb%218iBgEDuiKkK0A6OU5dxy3LpIE80YHMpIo7Jd6v3U1g3IdBXIfVtMRL1oFwNFwVkn%2Fitems%2F01QGC7QGHFA33C2EXFVBFZNQBRTUMELKX7%3Fversion%3DPublished%26access_token%3DeyJ0eXAiOiJKV1QiLCJhbGciOiJub25lIn0.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvYmRvY29sbGFiLnNoYXJlcG9pbnQuY29tQDVhMGE0OTc1LWYyYjMtNDRlNS1iYTA0LWQwMjkwOGFhNzU3NSIsImlzcyI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMCIsIm5iZiI6IjE2NDI3NzcyMDAiLCJleHAiOiIxNjQyNzk4ODAwIiwiZW5kcG9pbnR1cmwiOiIvQXJVSWxlSlFzSmxPQ0RkZStLSk9UZ3VsbUlzZXE4YkU0aXViTHVoT2RjPSIsImVuZHBvaW50dXJsTGVuZ3RoIjoiMTE2IiwiaXNsb29wYmFjayI6IlRydWUiLCJ2ZXIiOiJoYXNoZWRwcm9vZnRva2VuIiwic2l0ZWlkIjoiTVRBMk1ESXdaakl0WVRJellpMDBNakpoTFdJME1ETXRZVE01TkdVMVpHTTNNbVJqIiwiYXBwX2Rpc3BsYXluYW1lIjoiQXBwIFNlcnZpY2UiLCJnaXZlbl9uYW1lIjoiSHVtZXJhIiwiZmFtaWx5X25hbWUiOiJLaGFuIiwibmFtZWlkIjoiMCMuZnxtZW1iZXJzaGlwfGh1bWVyYWtAYmRvY29sbGFiLm9ubWljcm9zb2Z0LmNvbSIsIm5paSI6Im1pY3Jvc29mdC5zaGFyZXBvaW50IiwiaXN1c2VyIjoidHJ1ZSIsImNhY2hla2V5IjoiMGguZnxtZW1iZXJzaGlwfDEwMDMyMDAxYTgyN2RlZmVAbGl2ZS5jb20iLCJzaGFyaW5naWQiOiIxZDAzZmU1YS02YjEyLTRkMDMtYjM1NS0yOTczMzU3MjM0YzgiLCJ0dCI6IjAiLCJ1c2VQZXJzaXN0ZW50Q29va2llIjpudWxsLCJpcGFkZHIiOiI1Mi4yMzcuMjQuMTI2In0.Ty9IWFVoNUVSVWFsdkx0ajhMWUlrQ3lkeW9Hd1AzSVk2MzZZVytFWE1Mdz0%22%2C%22isFolder%22%3Atrue%7D%5D%7D&oAuthToken=");
      xhr.send("zipFileName=2021_2022-01-21_043553.zip&guid=7d13046a-2e64-4989-b71c-2ab2f80c93e7&provider=spo&files="+result+"&oAuthToken=");


    
  }
 });

} 

function FormatArrayWithResult(item, accessToken)
{
 var files =`{"items":[{"name":"`+item["FileLeafRef"]+`","size":"`+item["SMTotalSize"]+`","docId":"`+encodeURI(item[".spItemUrl"])+`&`+encodeURI(accessToken)+`","isFolder": true}]}`;
  return encodeURIComponent(files).replace(/[:]/g,"%3A").replace(/[!]/g,"%21");
 
}

function sortDesc(a, b)
{
var x = a["FileLeafRef"];
var y = b["FileLeafRef"];
return ((x > y) ? -1 : ((x < y) ? 1 : 0));
}








