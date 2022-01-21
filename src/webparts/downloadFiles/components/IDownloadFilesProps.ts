import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface IDownloadFilesProps {
  context: WebPartContext;
  siteUrl:string;
  ImageUrl:string;
  PartnerList:string;
  PartnerTaxFilesList:string;

  
}


