import { SPHttpClient } from '@microsoft/sp-http';

export interface IWPprojectdetailsProps {
  description: string;
  spHttpClient: SPHttpClient;
  siteurl: string;
  EmployeeName: string;
  EmployeeNumber: string;
  EmployeeEmail: string;
  ProjectName: string;
  ProjetDescription: string;
  ProjectManager: string;
  ProjectTeam: string;
  showPanel: boolean;
  _items:Array<object>;
  SelectedItemId:number;
  SelectedItemArray:Array<object>;
  AuthorDisplayName:string;
  BusinessManagerName:string;
  TeamMembersName:string;
  UserArray:Array<object>;
  UserIds:Array<string>;

}

//_items:any[];