import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
export interface IMovementProps {
  context: WebPartContext;
  pageSize?: number;
  viewType: string;
  users: any;
}
