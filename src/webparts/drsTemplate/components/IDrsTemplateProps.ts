import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IListService } from "../../services/SPListServices";
export interface IDrsTemplateProps {
  description: string;
  Title?: string;
  FlowURL?: string;
  LibraryId?: string;
  context?: WebPartContext;
  listServices: IListService;
  ProjectListId?: string;
  DRSApprovalsListId?:string;
  folderPath?:string;
  Filter:string;
}
