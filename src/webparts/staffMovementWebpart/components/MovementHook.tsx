import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './StaffMovementWebpart.module.scss';
import { PersonaCard } from "./PersonaCard/PersonaCard";
import {
  Spinner, SpinnerSize, MessageBar, MessageBarType, Label,
} from "office-ui-fabric-react";
import { initializeIcons } from '@fluentui/font-icons-mdl2';
import { Icon } from '@fluentui/react/lib/Icon';

import { Stack, IStackStyles, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import * as strings from "StaffMovementWebpartWebPartStrings";
import { IMovementState } from "./IMovementState";
import { IMovementProps } from './IMovementProps';
import * as moment from 'moment';
import Paging from './Pagination/Paging';
import { property } from 'lodash';

const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');
const wrapStackTokens: IStackTokens = { childrenGap: 30 };

const MovementHook: React.FC<IMovementProps> = (props) => {
  const [state, setstate] = useState<IMovementState>({
    users: [],
    isLoading: false,
    errorMessage: "",
    hasError: false,
  });
  const color = props.context.microsoftTeams ? "white" : "";
  const [pagedItems, setPagedItems] = useState<any[]>([]);
  const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
  const [currentPage, setCurrentPage] = useState<number>(1);

  const _onPageUpdate = async (pageno?: number) => {
    var currentPge = (pageno) ? pageno : currentPage;
    var startItem = ((currentPge - 1) * pageSize);
    var endItem = currentPge * pageSize;
    let filItems = slice(state.users, startItem, endItem);
    setCurrentPage(currentPge);
    setPagedItems(filItems);
  };

  initializeIcons();
  useEffect(() => {
    setPageSize(props.pageSize);
    if (state.users) _onPageUpdate();
  }, [state.users, props.pageSize]);
  if (props.viewType == "New") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          if (typeof user.Reporting_x0020_Officer === 'object') {
            var ro = true;
          }
          else ro = false;
          if (props.viewType == "New" && user.Status == "New") {
            var diff = moment(user.Join_x0020_Date).diff(moment(new Date), 'days');
            console.log(diff);
            if (diff > 6) {
              return null;
            }
            else {
              return (
                <PersonaCard
                  context={props.context}
                  profileProperties={{
                    DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                    Title: user.Designation ? user.Designation : '',
                    PictureUrl: user.PictureURL ? user.PictureURL : user.PictureURL,
                    Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                    Department: user.Department ? user.Department : user.Department,
                    WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                    MobilePhone: user.Mobile_x0020_Number
                      ? user.Mobile_x0020_Number
                      : '',
                    JoinDate: user.Join_x0020_Date && moment(user.Join_x0020_Date).isValid() ? moment(user.Join_x0020_Date).format("DD/MM/YYYY") : '',
                    ReportingOfficer: user.Reporting_x0020_Officer && ro ? user.Reporting_x0020_Officer.Title : ''
                  }}
                  viewType={props.viewType}
                />
              );
            }
          }
        })
        : [];

  }
  if (props.archivalType == "New-Archive") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          console.log(user.Name);
          if (typeof user.Reporting_x0020_Officer === 'object') {
            var ro = true;
          }
          else ro = false;
          if ((user.Status == "Active")) { //New-Archive

            var diff = moment(user.Join_x0020_Date).diff(moment(new Date), 'days');
            console.log(diff);
            if (diff >= -90 && diff <= -7) {
              return (
                <PersonaCard
                  context={props.context}
                  profileProperties={{
                    DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                    Title: user.Designation ? user.Designation : '',
                    PictureUrl: user.PictureURL ? user.PictureURL : user.PictureURL,
                    Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                    Department: user.Department ? user.Department : user.Department,
                    WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                    MobilePhone: user.Mobile_x0020_Number
                      ? user.Mobile_x0020_Number
                      : '',
                    JoinDate: user.Join_x0020_Date && moment(user.Join_x0020_Date).isValid() ? moment(user.Join_x0020_Date).format("DD/MM/YYYY") : '',
                    ReportingOfficer: user.Reporting_x0020_Officer && ro ? user.Reporting_x0020_Officer.Title : ''
                  }}
                  viewType={props.viewType}
                />
              );

            }
            else {
              return null;
            }
          }
        })
        : [];

  }

  if (props.viewType == "Transfer") { //test transfer tmr
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          if (typeof user.Reporting_x0020_Officer === 'object') {
            var ro = true;
          }
          else ro = false;
          if (props.viewType == "Transfer" && user.Status == "Transfer") {
            var diff = moment(user.Transfer_x0020_Date).diff(moment(new Date), 'days');
            if (diff > 6) {
              return null;
            }
            else {
              return (
                <PersonaCard
                  context={props.context}
                  profileProperties={{
                    DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                    Title: user.Designation ? user.Designation : '',
                    PictureUrl: user.PictureURL ? user.PictureURL : '',
                    Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                    Department: user.Department ? user.Department : '',
                    WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                    MobilePhone: user.Mobile_x0020_Number
                      ? user.Mobile_x0020_Number
                      : '',
                    TransferDate: user.Transfer_x0020_Date && moment(user.Transfer_x0020_Date).isValid() ? moment(user.Transfer_x0020_Date).format("DD/MM/YYYY") : '',
                    OldDepartment: user.OldDepartment ? user.OldDepartment : '',
                    OldDesignation: user.OldDepartment ? user.OldDesignation : '',
                    ReportingOfficer: user.Reporting_x0020_Officer && ro ? user.Reporting_x0020_Officer.Title : ''
                  }}
                  viewType={props.viewType}
                />
              );
            }
          }
        })
        : [];
  }

  if (props.archivalType == "Transfer-Archive") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          console.log(user.Name);
          if (typeof user.Reporting_x0020_Officer === 'object') {
            var ro = true;
          }
          else ro = false;
          if ((user.Status == "Active")) { //New-Archive

            var diff = moment(user.Transfer_x0020_Date).diff(moment(new Date), 'days');
            console.log(diff);
            if (diff >= -90 && diff <= -7) {
              return (
                <PersonaCard
                  context={props.context}
                  profileProperties={{
                    DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                    Title: user.Designation ? user.Designation : '',
                    PictureUrl: user.PictureURL ? user.PictureURL : '',
                    Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                    Department: user.Department ? user.Department : '',
                    WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                    MobilePhone: user.Mobile_x0020_Number
                      ? user.Mobile_x0020_Number
                      : '',
                    TransferDate: user.Transfer_x0020_Date && moment(user.Transfer_x0020_Date).isValid() ? moment(user.Transfer_x0020_Date).format("DD/MM/YYYY") : '',
                    OldDepartment: user.OldDepartment ? user.OldDepartment : '',
                    OldDesignation: user.OldDepartment ? user.OldDesignation : '',
                    ReportingOfficer: user.Reporting_x0020_Officer && ro ? user.Reporting_x0020_Officer.Title : ''
                  }}
                  viewType={props.viewType}
                />
              );

            }
            else {
              return null;
            }
          }
        })
        : [];

  }
  if (props.viewType == "Farewell") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          if (props.viewType == "Farewell" && user.Status == "Resigned") {
            var diff = moment(user.Last_x0020_Service_x0020_Date).diff(moment(new Date), 'days');
            console.log(diff);
            if (diff > 13) {
              return null;
            }
            else {
              return (

                <PersonaCard
                  context={props.context}
                  profileProperties={{
                    DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                    Title: user.Designation ? user.Designation : '',
                    PictureUrl: user.PictureURL ? user.PictureURL : '',
                    Department: user.Department ? user.Department : '',
                    Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                    WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                    MobilePhone: user.Mobile_x0020_Number
                      ? user.Mobile_x0020_Number
                      : '',
                    LastServiceDate: user.Last_x0020_Service_x0020_Date && moment(user.Last_x0020_Service_x0020_Date).isValid() ? moment(user.Last_x0020_Service_x0020_Date).format("DD/MM/YYYY") : '',
                  }}
                  viewType={props.viewType}
                />
              );
            }
          }
        })
        : [];
  }

  if (props.archivalType == "Farewell-Archive") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          console.log(user.Name);
          if (typeof user.Reporting_x0020_Officer === 'object') {
            var ro = true;
          }
          else ro = false;
          if ((user.Status == "Inactive")) { //New-Archive

            var diff = moment(user.Last_x0020_Service_x0020_Date).diff(moment(new Date), 'days');
            console.log(diff);
            if (diff >= -30 && diff <= -7) {
              return (
                <PersonaCard
                  context={props.context}
                  profileProperties={{
                    DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                    Title: user.Designation ? user.Designation : '',
                    PictureUrl: user.PictureURL ? user.PictureURL : '',
                    Department: user.Department ? user.Department : '',
                    Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                    WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                    MobilePhone: user.Mobile_x0020_Number
                      ? user.Mobile_x0020_Number
                      : '',
                    LastServiceDate: user.Last_x0020_Service_x0020_Date && moment(user.Last_x0020_Service_x0020_Date).isValid() ? moment(user.Last_x0020_Service_x0020_Date).format("DD/MM/YYYY") : '',
                  }}
                  viewType={props.viewType}
                />
              );

            }
            else {
              return null;
            }
          }
        })
        : [];

  }

  return (
    <div className={styles.directory}>
      {state.isLoading ? (
        <div style={{ marginTop: '10px' }}>
          <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
        </div>
      ) : (
        <>
          {state.hasError ? (
            <div style={{ marginTop: '10px' }}>
              <MessageBar messageBarType={MessageBarType.error}>
                {state.errorMessage}
              </MessageBar>
            </div>
          ) : (
            <>
              {!props.users || props.users == 0 ? (
                <div className={styles.noUsers}>
                  <Icon
                    iconName={"ProfileSearch"}
                    style={{ fontSize: "54px", color: color }}
                  />
                  <Label>
                    <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                      {strings.DirectoryMessage}
                    </span>
                  </Label>
                </div>
              ) : (
                <>
                  {/* <div style={{ width: '100%', display: 'inline-block' }}>
                                    <Paging
                                        totalItems={state.users.length}
                                        itemsCountPerPage={pageSize}
                                        onPageUpdate={_onPageUpdate}
                                        currentPage={currentPage} />
                                </div> */}

                  {/* <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens} style={{ width: "100em" }}> */}
                  <Stack horizontalAlign="center" wrap tokens={wrapStackTokens} style={{ width: "100em", maxWidth: "100em" }}>
                    <div style={{ marginTop: "2em", marginLeft: "3em" }}>{movementGrid}</div>
                  </Stack>
                  {/* <div style={{ width: '100%', display: 'inline-block' }}>
                    <Paging
                      totalItems={state.users.length}
                      itemsCountPerPage={pageSize}
                      onPageUpdate={_onPageUpdate}
                      currentPage={currentPage} />
                  </div> */}
                </>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
};
export default MovementHook;
