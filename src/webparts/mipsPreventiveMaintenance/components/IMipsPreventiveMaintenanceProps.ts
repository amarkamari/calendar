import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IMipsPreventiveMaintenanceProps {
  siteUrl: string;
  listName: string;
  allowEdit: boolean;
  context: WebPartContext;
}