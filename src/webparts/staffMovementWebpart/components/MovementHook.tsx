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
          return (
            <PersonaCard
              context={props.context}
              profileProperties={{
                DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                Title: user.Designation && user.Designation.JobTitle ? user.Designation.JobTitle : '',
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
            />
          );
        })
        : [];
  }
  if (props.viewType == "Transfer") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          if (typeof user.Reporting_x0020_Officer === 'object') {
            var ro = true;
          }
          else ro = false;
          return (
            <PersonaCard
              context={props.context}
              profileProperties={{
                DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                Title: user.Designation && user.Designation.JobTitle ? user.Designation.JobTitle : '',
                PictureUrl: user.PictureURL ? user.PictureURL : '',
                Email: user.Email_x0020_Address && user.Email_x0020_Address.EMail ? user.Email_x0020_Address.EMail : '',
                Department: user.Department ? user.Department : '',
                WorkPhone: user.DID && user.DID.WorkPhone ? user.DID.WorkPhone : '',
                MobilePhone: user.Mobile_x0020_Number
                  ? user.Mobile_x0020_Number
                  : '',
                TransferDate: user.Transfer_x0020_Date && moment(user.Transfer_x0020_Date).isValid() ? moment(user.Transfer_x0020_Date).format("DD/MM/YYYY") : '',
                OldDepartment: user.OldDepartment ? user.OldDepartment : '',
                ReportingOfficer: user.Reporting_x0020_Officer && ro ? user.Reporting_x0020_Officer.Title : ''
              }}
            />
          );
        })
        : [];
  }
  if (props.viewType == "Farewell") {
    var movementGrid =
      props.users && props.users.length > 0
        ? props.users.map((user: any) => {
          return (
            <PersonaCard
              context={props.context}
              profileProperties={{
                DisplayName: user.Name && user.Name.Title ? user.Name.Title : '',
                Title: user.Designation && user.Designation.JobTitle ? user.Designation.JobTitle : '',
                PictureUrl: user.PictureURL ? user.PictureURL : '',
                Department: user.Department ? user.Department : '',
                LastServiceDate: user.Last_x0020_Service_x0020_Date && moment(user.Last_x0020_Service_x0020_Date).isValid() ? moment(user.Last_x0020_Service_x0020_Date).format("DD/MM/YYYY") : '',
              }}
            />
          );
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

                  <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
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
