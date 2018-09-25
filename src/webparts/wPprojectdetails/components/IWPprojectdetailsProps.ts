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
  _items: Array<object>;
  SelectedItemId: number;
  SelectedItemArray: Array<object>;
  AuthorDisplayName: string;
  BusinessManagerName: string;
  TeamMembersName: string;
  UserArray: Array<object>;
  UserIds: Array<string>;
  PersonaArray: Array<object>;
  PersonNameArray: Array<string>;

  ProjectOpenAirId: string;
  ProjectCode: string;
  Client: string;
  ClientContact: string;
  ProjectStartDate: string;
  ProjectPlannedHours: string;
  ProjectPlannedDays: string;
  ProjectReportingRequirements: string;
  ProjectDeliverables: string;
  ProjectUsefulLinks: string;

}

//_items:any[];