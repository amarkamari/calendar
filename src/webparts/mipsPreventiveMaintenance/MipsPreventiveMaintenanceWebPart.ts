import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MipsPreventiveMaintenanceWebPartStrings';
import MipsPreventiveMaintenance from './components/MipsPreventiveMaintenance';
import { IMipsPreventiveMaintenanceProps } from './components/IMipsPreventiveMaintenanceProps';

export interface IMipsPreventiveMaintenanceWebPartProps {
  listName: string;
  allowEdit: boolean;
}

export default class MipsPreventiveMaintenanceWebPart
  extends BaseClientSideWebPart<IMipsPreventiveMaintenanceWebPartProps> {

  public render(): void {

const element: React.ReactElement<IMipsPreventiveMaintenanceProps> =
  React.createElement(
    MipsPreventiveMaintenance,
    {
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listName: this.properties.listName || 'PM_Masterdata_Test',
      allowEdit: this.properties.allowEdit !== false,
      context: this.context
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
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name'
                }),

                PropertyPaneToggle('allowEdit', {
                  label: 'Allow editing maintenance dates',
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
