import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'StaffMovementWebpartWebPartStrings';
import StaffMovementWebpart from './components/MovementHook';
import { IStaffMovementWebpartProps } from './components/IStaffMovementWebpartProps';
import { sp } from '@pnp/sp';
import { TooltipHost } from 'office-ui-fabric-react';
import MovementHook from "./components/MovementHook";
import { IMovementProps } from "./components/IMovementProps";

export interface IStaffMovementWebpartWebPartProps {
  listName: string;
  viewType: string;
  archivalType: string;
  pageSize: number;
}

export default class StaffMovementWebpartWebPart extends BaseClientSideWebPart<IStaffMovementWebpartWebPartProps> {

  public async render() {
    sp.setup({
      spfxContext: this.context
    });
    let items;
    items = await this.getListItems(this.properties.viewType);


    const element: React.ReactElement<IMovementProps> = React.createElement(
      MovementHook,
      {
        context: this.context,
        pageSize: this.properties.pageSize,
        viewType: this.properties.viewType,
        archivalType: this.properties.archivalType,
        users: items
      }
    );

    ReactDom.render(element, this.domElement);
  }
  //get items from the list based on the viewtype NOTE: Change value for query
  public async getListItems(viewType) {
    if (viewType == 'New') {
      const users: any[] = await sp.web.lists.getByTitle(this.properties.listName).items.select('Name/Title', 'Designation', 'OldDesignation', 'Email_x0020_Address/EMail', 'DID/WorkPhone', 'Mobile_x0020_Number', 'Department', 'Join_x0020_Date', 'Status', 'Reporting_x0020_Officer/Title').expand('Name', 'Email_x0020_Address', 'DID', 'Reporting_x0020_Officer').filter(`Status eq 'New' or Status eq 'Active'`).orderBy('Join_x0020_Date', false).get();
      if (users.length > 0) {
        for (let index = 0; index < users.length; index++) {
          let user: any = users[index];
          if (user && user.Email_x0020_Address) {
            user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.Email_x0020_Address.EMail}` };
            users[index] = user;
          }
        }
      }
      console.log(users);
      return users;
    }
    else if (viewType == 'Transfer') {
      const users: any[] = await sp.web.lists.getByTitle(this.properties.listName).items.select('Name/Title', 'Designation', 'OldDesignation', 'Email_x0020_Address/EMail', 'DID/WorkPhone', 'Mobile_x0020_Number', 'Department', 'OldDepartment', 'Transfer_x0020_Date', 'Status', 'Reporting_x0020_Officer/Title').expand('Name', 'Email_x0020_Address', 'DID', 'Reporting_x0020_Officer').filter(`Status eq 'Transfer' or Status eq 'Active'`).orderBy('Transfer_x0020_Date', false).get();
      if (users.length > 0) {
        for (let index = 0; index < users.length; index++) {
          let user: any = users[index];
          if (user && user.Email_x0020_Address) {
            user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.Email_x0020_Address.EMail}` };
            users[index] = user;
          }
        }
        return users;
      }
    }
    else if (viewType == 'Farewell') {
      const users: any[] = await sp.web.lists.getByTitle(this.properties.listName).items.select('Name/Title', 'Designation', 'OldDesignation', 'Email_x0020_Address/EMail', 'DID/WorkPhone', 'Mobile_x0020_Number', 'Department', 'Last_x0020_Service_x0020_Date', 'Status').expand('Name', 'Email_x0020_Address', 'DID').filter(`Status eq 'Resigned' or Status eq 'Inactive'`).orderBy('Last_x0020_Service_x0020_Date', false).get();
      if (users.length > 0) {
        for (let index = 0; index < users.length; index++) {
          let user: any = users[index];
          if (user && user.Email_x0020_Address) {
            user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.Email_x0020_Address.EMail}` };
            users[index] = user;
          }
        }
        return users;
      }
    }
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
                  label: "List Name"
                }),
                PropertyPaneDropdown('viewType', {
                  label: 'View Type',
                  options: [
                    { key: 'New', text: 'New' },
                    { key: 'Transfer', text: 'Transfer' },
                    { key: 'Farewell', text: 'Farewell' },
                  ]
                }),
                PropertyPaneDropdown('archivalType', {
                  label: 'Archival Type',
                  options: [
                    { key: 'New-Archive', text: 'New-Archive' },
                    { key: 'Transfer-Archive', text: 'Transfer-Archive' },
                    { key: 'Farewell-Archive', text: 'Farewell-Archive' }
                  ]
                }),
                PropertyPaneSlider('pageSize', {
                  label: 'Results per page',
                  showValue: true,
                  max: 20,
                  min: 2,
                  step: 2,
                  value: this.properties.pageSize
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
