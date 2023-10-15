import { IListService } from "../../../services/SPListServices";

import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDashboardProps {
    DRSLibraryID?: string;
    listServices?: IListService;
    Filter?: string;
    displayItemsCount?: number;
    CurrentUser?: any;
    isMyRequests?: boolean;
    folderPath?: string;
    context?: WebPartContext;
    DRSApprovalsListId?:string;
}
