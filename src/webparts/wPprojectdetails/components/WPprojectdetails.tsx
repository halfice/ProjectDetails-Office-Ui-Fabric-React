import * as React from 'react';
import styles from './WPprojectdetails.module.scss';
import { IWPprojectdetailsProps } from './IWPprojectdetailsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { default as pnp, ItemAddResult, Web, ReorderingRuleMatchType, RoleDefinitionBindings } from "sp-pnp-js";

import { DefaultButton } from 'office-ui-fabric-react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import 'office-ui-fabric-react/lib/components/List/Examples/List.Ghosting.Example.scss';

import { IPersonaSharedProps, Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Label } from 'office-ui-fabric-react/lib/Label';

const examplePersona: IPersonaSharedProps = {
  imageInitials: 'AL',
  text: 'Annie Lindqvist',
  secondaryText: 'Software Engineer',
  tertiaryText: 'In a meeting',
  optionalText: 'Available at 4:00pm'
};




export default class WPprojectdetails extends React.Component<IWPprojectdetailsProps, {}> {

  public state: IWPprojectdetailsProps;

  constructor(props, context) {
    super(props);
    this.state = {
      spHttpClient: this.props.spHttpClient,
      description: "",
      siteurl: this.props.siteurl,
      EmployeeName: "",
      EmployeeNumber: "",
      EmployeeEmail: "",
      ProjectName: "",
      ProjetDescription: "",
      ProjectManager: "",
      ProjectTeam: "",
      _items: [],
      showPanel: false,
      SelectedItemId: 0,
      SelectedItemArray: [],
      AuthorDisplayName: "",
      BusinessManagerName: "",
      TeamMembersName: "",
      UserArray: [],
      UserIds: [],

    };
    this.itemdetails = this.itemdetails.bind(this);
    this.findNext = this.findNext.bind(this);
    this.findPrevious = this.findPrevious.bind(this);
  }




  componentDidMount() {
    this.GetUSerDetails();
  }


  private GetUSerDetails() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    var TempEmailFinal = "";
    pnp.sp.profiles.myProperties.get().then(function (result) {
      var props = result.UserProfileProperties;
      var propValue = "";

      props.forEach(function (prop) {
        if (prop.Key == "AccountName") {
          var TempEmail = prop.Value;
          var TempEmailFinal = TempEmail.replace("i:0#.f|membership|", "");
        }
      });
    });
    this.getProjectdetails();
    this.setState({
      Loading: 0,
      EmployeeEmail: TempEmailFinal
    });



  }



  public ParsingUserNames() {
    this.getuserNamesById();
  }

  public getuserNamesById() {

    var TempUserArray = [];
    var TmpArray = this.state.UserIds;
    for (var i = 0; i < this.state.UserIds.length; i++) {
      var NewISiteUrl = this.props.siteurl;
      var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
      let webx = new Web(NewSiteUrl);
      var Tmpstr = "";
      Tmpstr = TmpArray[i];
      var NumberParse = parseInt(Tmpstr);
      webx.siteUsers.getById(NumberParse).get().then(function (result) {
        var NewItem = {
          Id: result.Id,
          Name: result.Title,
        }
        TempUserArray.push(NewItem);
      });
    }
    this.setState({
      UserArray: TempUserArray
    });




  }


  public inArray(needle, haystack) {
    var length = haystack.length;
    for (var i = 0; i < length; i++) {
      if (haystack[i] == needle)
        return true;
    }
    return false;
  }


  private getProjectdetails() {
    var NewISiteUrl = this.props.siteurl;
    var NewSiteUrl = NewISiteUrl.replace("/SitePages", "");
    let webx = new Web(NewSiteUrl);
    webx.lists.getByTitle('Demo2').items.get().then(item => {
      console.log(item);
      var ArTemp = [];
      var UsrTemp = [];
      var Counter = 1;
      for (var i = 0; i < item.length; i++) {
        var NewData = {
          key: Counter,
          thumbnail: item[i].Title,
          cover: item[i].Title,
          name: item[i].Title,
          description: item[i].Description,
          index: Counter,
          id: item[i].Id,
          AuthoName: item[i].AuthorId,
          ManagerName: item[i].ManagerId,
          TeamMemebers: item[i].TeamMembersId,
        }
        ArTemp.push(NewData);
        if (this.inArray(item[i].AuthorId, UsrTemp) == false) {
          UsrTemp.push(item[i].AuthorId);
        }
        if (this.inArray(item[i].ManagerId, UsrTemp) == false) {
          UsrTemp.push(item[i].ManagerId);
        }
        if (item[i].TeamMembersId != null && item[i].TeamMembersId.length > 0) {
          var tmp = item[i].TeamMembersId;
          for (var x = 0; x < tmp.length; x++) {
            if (this.inArray(tmp[x], UsrTemp) == false) {
              UsrTemp.push(tmp[x]);
            }
          }
        } else {
          if (this.inArray(item[i].TeamMembersId, UsrTemp) == false) {
            UsrTemp.push(item[i].TeamMembersId);
          }
        }

      }
      this.setState({
        _items: ArTemp,
        UserIds: UsrTemp,
      });
      this.ParsingUserNames();

    });
  }

  private _setShowPanel = (showPanel: boolean): (() => void) => {
    return (): void => {
      this.setState({ showPanel });
    };
  };

  public itemdetails(event: any): void {
    var tmpString = event.target.id;
    var tmpid = parseInt(tmpString);
    var TempAllArray = this.state._items;
    var TempComplete = [];

    TempComplete = TempAllArray.filter(function (TempAllArray) {
      return TempAllArray["id"] == tmpid;
    });


    var AuthorIdInArray = TempComplete[0]["AuthoName"];
    var ManagerIdInArray = TempComplete[0]["ManagerName"];
    var TeammembersIdsArray = TempComplete[0]["TeamMemebers"];
    var UserArray1 = this.state.UserArray;
    UserArray1 = UserArray1.filter(function (UserArray1) {
      return UserArray1["Id"] == AuthorIdInArray;
    });
    TempComplete[0]["AuthoName"] = UserArray1[0]["Name"];


    UserArray1 = this.state.UserArray;
    UserArray1 = UserArray1.filter(function (UserArray1) {
      return UserArray1["Id"] == ManagerIdInArray;
    });
    TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];

    var TeamMemebrNameString = [];
    for (var y = 0; y < TeammembersIdsArray.length; y++) {
      var UserArray1 = this.state.UserArray;
      UserArray1 = UserArray1.filter(function (UserArray1) {
        return UserArray1["Id"] == TeammembersIdsArray[y];
      });
      TeamMemebrNameString.push(UserArray1[0]["Name"]);
    }
    TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;



    this.setState(
      {
        SelectedItemId: parseInt(tmpString),
        showPanel: true,
        SelectedItemArray: TempComplete,
      }
    )
  }




  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  };

  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <DefaultButton onClick={this._onClosePanel} style={{ marginRight: '8px' }}>
          Save
        </DefaultButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  };

  public findNext() {
    var tempCurrenItem = this.state.SelectedItemId;
    if (tempCurrenItem < this.state._items.length) {
      var NewTempNumber = tempCurrenItem + 1;
      var TempAllArray = this.state._items;
      var TempComplete = TempAllArray.filter(function (TempAllArray) {
        return TempAllArray["id"] == NewTempNumber;
      });

      var AuthorIdInArray = TempComplete[0]["AuthoName"];
      var ManagerIdInArray = TempComplete[0]["ManagerName"];
      var TeammembersIdsArray = TempComplete[0]["TeamMemebers"];

      if (isNaN(AuthorIdInArray) == false) {
        var UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == AuthorIdInArray;
        });
        TempComplete[0]["AuthoName"] = UserArray1[0]["Name"];
      }

      if (isNaN(ManagerIdInArray) == false) {
        UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == ManagerIdInArray;
        });
        TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];
      }

      var TeamMemebrNameString = [];

      var conversioncounter = 0;
      for (var y = 0; y < TeammembersIdsArray.length; y++) {
        if (isNaN(TeammembersIdsArray[y]) == false) {
          var UserArray1 = this.state.UserArray;
          UserArray1 = UserArray1.filter(function (UserArray1) {
            return UserArray1["Id"] == TeammembersIdsArray[y];
          });
          TeamMemebrNameString.push(UserArray1[0]["Name"]);
          conversioncounter++;
        }
      }

      if (conversioncounter > 0) {
        TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;
      }

      this.setState(
        {
          SelectedItemId: NewTempNumber,
          showPanel: true,
          SelectedItemArray: TempComplete,
        });
    }


  }

  public findPrevious() {
    var tempCurrenItem = this.state.SelectedItemId;
    var NewTempNumber = tempCurrenItem - 1;
    if (NewTempNumber >= 1) {
      var TempAllArray = this.state._items;
      var TempComplete = TempAllArray.filter(function (TempAllArray) {
        return TempAllArray["id"] == NewTempNumber;
      });

      var AuthorIdInArray = TempComplete[0]["AuthoName"];
      var ManagerIdInArray = TempComplete[0]["ManagerName"];
      var TeammembersIdsArray = TempComplete[0]["TeamMemebers"];



      if (isNaN(AuthorIdInArray) == false) {
        var UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == AuthorIdInArray;
        });
        TempComplete[0]["AuthoName"] = UserArray1[0]["Name"];
      }

      if (isNaN(ManagerIdInArray) == false) {
        UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == ManagerIdInArray;
        });
        TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];
      }

      var TeamMemebrNameString = [];

      var conversioncounter = 0;
      for (var y = 0; y < TeammembersIdsArray.length; y++) {
        if (isNaN(TeammembersIdsArray[y]) == false) {
          var UserArray1 = this.state.UserArray;
          UserArray1 = UserArray1.filter(function (UserArray1) {
            return UserArray1["Id"] == TeammembersIdsArray[y];
          });
          TeamMemebrNameString.push(UserArray1[0]["Name"]);
          conversioncounter++;
        }
      }

      if (conversioncounter > 0) {
        TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;
      }




      this.setState(
        {
          SelectedItemId: NewTempNumber,
          showPanel: true,
          SelectedItemArray: TempComplete,
        });

    }
  }




  public render(): React.ReactElement<IWPprojectdetailsProps> {

    const alertClicked = (): void => {
      alert('Clicked');
    };



    if (this.state._items.length > 0) {
      var FindingArrayOption = this.state._items.map(function (item, i) {
        return (<div >
          <div className="ms-ListGhostingExample-itemCell">
            <span className={styles.PrimaryItem}>{item["id"]}</span>
            <div className="ms-ListGhostingExample-itemContent">
              <div className="ms-ListGhostingExample-itemName" onClick={this.itemdetails.bind(this)} >
                <span className={styles.myspan} id={item["id"]}  >
                  {item["name"]}</span>
              </div>
              <div className="ms-ListGhostingExample-itemIndex">{item["description"]}</div>
            </div>
          </div>
        </div>)
      }, this);
    }



    return (
      <div className={styles.wPprojectdetails}>

        {
          FindingArrayOption
        }


        <Panel
          isOpen={this.state.showPanel}
          type={PanelType.medium}
          onDismiss={this._onClosePanel}
          isFooterAtBottom={true}
          headerText="Projects Information Tool"
          closeButtonAriaLabel="Close"
          onRenderFooterContent={this._onRenderFooterContent}
        >

          <DefaultButton
            style={{ color: 'Black', backgroundColor: 'white' }}
            text='All Projects'
            onClick={this._onClosePanel}
          />

          <DefaultButton
            style={{ color: 'Black', backgroundColor: 'white' }}
            text='Previous' onClick={this.findPrevious}
          />


          <DefaultButton
            style={{ color: 'Black', backgroundColor: 'white' }}
            text='Next'
            onClick={this.findNext}
          />


          <span>

            {

              this.state.SelectedItemArray.length > 0 &&
              <div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project Name</div>
                    <div className="ms-Grid-col ms-u-sm6 block">{this.state.SelectedItemArray[0]["name"]}</div>
                  </div>
                </div>


                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Author - Request By</div>
                    <div className="ms-Grid-col ms-u-sm6 block">{this.state.SelectedItemArray[0]["AuthoName"]}
                    <Persona
                      {...examplePersona}
                      size={PersonaSize.size24}
                      presence={PersonaPresence.online}
                      hidePersonaDetails={false}
                    />
                    </div>
                  </div>
                </div>



                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project Manager</div>
                    <div className="ms-Grid-col ms-u-sm6 block">
                    <Persona
                      {...examplePersona}
                      size={PersonaSize.size24}
                      presence={PersonaPresence.online}
                      hidePersonaDetails={false}
                    />
                    </div>
                   
                  </div>
                </div>




                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Team Members</div>
                    <div className="ms-Grid-col ms-u-sm6 block">{this.state.SelectedItemArray[0]["TeamMemebers"]}</div>
                  </div>
                </div>


                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Client Contact</div>
                    <div className="ms-Grid-col ms-u-sm6 block">+9134872984732984</div>
                  </div>
                </div>


                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Client Emails</div>
                    <div className="ms-Grid-col ms-u-sm6 block">client@onmicrosoft.com</div>
                  </div>
                </div>


                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project Code</div>
                    <div className="ms-Grid-col ms-u-sm6 block">{this.state.SelectedItemArray[0]["id"]}</div>
                  </div>
                </div>
              </div>




            }





          </span>



        </Panel>

      </div >
    );
  }






}
