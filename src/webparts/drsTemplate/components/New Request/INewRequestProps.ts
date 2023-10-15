import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IListService } from "../../../services/SPListServices";
export interface INewRequestProps {
    context?: WebPartContext;
    DRSLibray?: string;
    FlowURL?: string;
    listServices?: IListService;
    ProjectListId?: string;
    ProjectNumberListId?: string;

}

export interface ICarrierOptions {
    title: string;
}
