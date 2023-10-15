import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DrsTemplateWebPartStrings';
import DrsTemplate from './components/DrsTemplate';
import { IDrsTemplateProps } from './components/IDrsTemplateProps';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { get, update } from '@microsoft/sp-lodash-subset';
import { IListService, ListService } from "../services/SPListServices";

require('../CustomCss.css');
export interface IDrsTemplateWebPartProps {
  description: string;
  ProjectListId: string;
  DRSApprovalsListId: string;
  DRSLibrary: string;
  FlowURL: string;
  Title: string;
  folderPath: string;
  Filter: string;
}

export default class DrsTemplateWebPart extends BaseClientSideWebPart<IDrsTemplateWebPartProps> {
  private ListServices: IListService;

  protected onInit(): Promise<void> {
    this.ListServices = new ListService(this.context);

    return super.onInit();
  }
  public render(): void {
    const element: React.ReactElement<IDrsTemplateProps> = React.createElement(
      DrsTemplate,
      {
        description: this.properties.description,
        context: this.context,
        listServices: this.ListServices,
        FlowURL: this.properties.FlowURL,
        ProjectListId: this.properties.ProjectListId,
        DRSApprovalsListId: this.properties.DRSApprovalsListId,
        folderPath: this.properties.folderPath,
        Filter: this.properties.Filter
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private onCustomPropertyPaneChange(propertyPath: string, newValue: any): void {
    // Log.verbose(this.logSource, "WebPart property '" + propertyPath + "' has changed, refreshing WebPart...", this.context.serviceScope);
    const oldValue = get(this.properties, propertyPath);
    // Stores the new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('folderPath', {
                  label: "Server Relative URL for Library"
                }),
                PropertyPaneTextField('FlowURL', {
                  label: "Flow URL to trigger "
                }),
                PropertyPaneTextField('Filter', {
                  label: "Team Projects "
                }),
                PropertyFieldListPicker('ProjectListId', {
                  label: "Select Project List",
                  selectedList: this.properties.ProjectListId,
                  includeHidden: false,
                  baseTemplate: 100,
                  disabled: false,
                  onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerRequestFieldId'
                }),
                PropertyFieldListPicker('DRSApprovalsListId', {
                  label: "Select DRSApprovers List",
                  selectedList: this.properties.DRSApprovalsListId,
                  includeHidden: false,
                  baseTemplate: 100,
                  disabled: false,
                  onPropertyChange: this.onCustomPropertyPaneChange.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerRequestFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
