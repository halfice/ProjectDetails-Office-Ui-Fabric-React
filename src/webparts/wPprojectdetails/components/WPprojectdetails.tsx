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
import { HoverCard, IExpandingCardProps, ExpandingCardMode } from 'office-ui-fabric-react/lib/HoverCard';

import { DirectionalHint } from 'office-ui-fabric-react/lib/common/DirectionalHint';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';


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
      PersonaArray: [],
      PersonNameArray: [],

      ProjectOpenAirId: "",
      ProjectCode: "",
      Client: "",
      ClientContact: "",
      ProjectStartDate: "",
      ProjectPlannedHours: "",
      ProjectPlannedDays: "",
      ProjectReportingRequirements: "",
      ProjectDeliverables: "",
      ProjectUsefulLinks: "",
      CurrentIemCounter: 1,





    };
    this.itemdetails = this.itemdetails.bind(this);
    this.findNext = this.findNext.bind(this);
    this.findPrevious = this.findPrevious.bind(this);
  }

  componentDidMount() {

    if (Environment.type === EnvironmentType.Local) {
      var ArTemp = [];
      var UsrTemp = [];
      var Counter = 1;
      var NewData = {
        key: 1,
        thumbnail: "Title",
        cover: "Cover Title",
        name: "Project A",
        description: "Project Description",
        index: 1,
        id: 1,
        AuthoName: "1",
        ManagerName: "2",
        TeamMemebers: "3",
      }
      ArTemp.push(NewData);
      
      var NewData2 = {
        key: 2,
        thumbnail: "Title 1",
        cover: "Cover Title ",
        name: "Project B",
        description: "Project Description B",
        index: 2,
        id: 2,
        AuthoName: "2",
        ManagerName: "1",
        TeamMemebers: "3",
      }
      ArTemp.push(NewData2);
      UsrTemp.push(1);
      UsrTemp.push(2);
      UsrTemp.push(3);

      this.setState({
        _items: ArTemp,
        UserIds: UsrTemp,
      });

      var TempUserArray = [];

      var NewItem = {
        Id: 1,
        Name: "Test User 1",
        Organization: "Org A",
        Email: "test@gmail.com",
        Tel: "33333333",
        Location: "City A",
      }
      TempUserArray.push(NewItem);

      var NewItem2 = {
        Id: 2,
        Name: "Test User 2",
        Organization: "Org B",
        Email: "test2@gmail.com",
        Tel: "88888",
        Location: "City B",
      }
      TempUserArray.push(NewItem2);

      var NewItem3 = {
        Id: 3,
        Name: "Test User 3",
        Organization: "Org C",
        Email: "test3@gmail.com",
        Tel: "888888",
        Location: "City C",
      }
      TempUserArray.push(NewItem3);

      this.setState({
        Loading: 0,
        EmployeeEmail: 'test@onmicrosoft.com',
        _items: ArTemp,
        UserIds: UsrTemp,
        UserArray: TempUserArray
      });

    } else {
      this.GetUSerDetails();
    }


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
      EmployeeEmail: TempEmailFinal,

    });



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
        Counter++;
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
          Organization: result.Title,
          Email: result.Title,
          Tel: result.Title,
          Location: result.Title,
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
  private _setShowPanel = (showPanel: boolean): (() => void) => {
    return (): void => {
      this.setState({ showPanel });
    };
  };

  public itemdetails(event: any): void {
    if (Environment.type === EnvironmentType.Local) {
      this.itemdetailsDummy(event.target.id);
      return;
    }

    var tmpString = event.target.id;
    var tmpid = parseInt(tmpString);
    var TempAllArray = this.state._items;
    var TempComplete = [];

    var TempPerosonaArra = [];

    TempComplete = TempAllArray.filter(function (TempAllArray) {
      return TempAllArray["id"] == tmpid;
    });


    var AuthorIdInArray = TempComplete[0]["AuthoName"];
    var ManagerIdInArray = TempComplete[0]["ManagerName"];
    var TeammembersIdsArray = TempComplete[0]["TeamMemebers"];

    var FinalName = "";
    var UserArray1 = this.state.UserArray;
    if (isNaN(AuthorIdInArray) == false) {
      UserArray1 = UserArray1.filter(function (UserArray1) {
        return UserArray1["Id"] == AuthorIdInArray;
      });
      TempComplete[0]["AuthoName"] = UserArray1[0]["Name"];
      FinalName = UserArray1[0]["Name"];
    } else {
      FinalName = TempComplete[0]["AuthoName"];
    }

    var NewObject1 = {
      imageInitials: '',
      text: FinalName,
      secondaryText: 'Software Engineer',
      tertiaryText: 'In a meeting',
      optionalText: 'Available at 4:00pm'
    }
    TempPerosonaArra.push(NewObject1);



    UserArray1 = this.state.UserArray;
    if (isNaN(ManagerIdInArray) == false) {
      UserArray1 = UserArray1.filter(function (UserArray1) {
        return UserArray1["Id"] == ManagerIdInArray;
      });
      TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];
      FinalName = UserArray1[0]["Name"];
    } else {
      FinalName = TempComplete[0]["ManagerName"];
    }

    var NewObject2 = {
      imageInitials: '',
      text: FinalName,
      secondaryText: 'Software Engineer',
      tertiaryText: 'In a meeting',
      optionalText: 'Available at 4:00pm'
    }
    TempPerosonaArra.push(NewObject2);

    var Arobject = [];

    var TeamMemebrNameString = [];

    for (var y = 0; y < TeammembersIdsArray.length; y++) {

      var UserArray1 = this.state.UserArray;
      if (isNaN(TeammembersIdsArray[y]) == false) {
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == TeammembersIdsArray[y];
        });
        TeamMemebrNameString.push(UserArray1[0]["Name"]);
      } else {
        TeamMemebrNameString.push(TeammembersIdsArray[y]);
      }
    }
    TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;
    TempPerosonaArra.push(Arobject);
    this.setState(
      {
        SelectedItemId: parseInt(tmpString),
        showPanel: true,
        SelectedItemArray: TempComplete,
        PersonaArray: TempPerosonaArra,
        PersonNameArray: TeamMemebrNameString,
        CurrentIemCounter: TempComplete[0]["key"],
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
    var tempCurrenItem = this.state.CurrentIemCounter;
    var TempPerosonaArra = [];
    var FinalName = "";

    if (tempCurrenItem < this.state._items.length) {
      var NewTempNumber = tempCurrenItem + 1;
      var TempAllArray = this.state._items;
      var TempComplete = TempAllArray.filter(function (TempAllArray) {
        return TempAllArray["key"] == NewTempNumber;
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
        FinalName = UserArray1[0]["Name"];
      } else {
        FinalName = TempComplete[0]["AuthoName"];

      }
      var NewObject1 = {
        imageInitials: '',
        text: FinalName,
        secondaryText: 'Software Engineer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      }
      TempPerosonaArra.push(NewObject1);


      if (isNaN(ManagerIdInArray) == false) {
        UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == ManagerIdInArray;
        });
        TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];
        FinalName = UserArray1[0]["Name"];
      } else {
        FinalName = TempComplete[0]["ManagerName"];
      }

      var NewObject2 = {
        imageInitials: '',
        text: FinalName,
        secondaryText: 'Software Engineer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      }
      TempPerosonaArra.push(NewObject2);

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
        } else {
          TeamMemebrNameString.push(TeammembersIdsArray[y]);
        }
      }

      if (conversioncounter > 0) {
        TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;
      }


      this.setState(
        {

          CurrentIemCounter: NewTempNumber,
          SelectedItemId: TempComplete[0]["Id"],
          showPanel: true,
          SelectedItemArray: TempComplete,
          PersonaArray: TempPerosonaArra,
          PersonNameArray: TeamMemebrNameString,

        });
    }


  }

  public findPrevious() {
    var tempCurrenItem = this.state.CurrentIemCounter;
    var NewTempNumber = tempCurrenItem - 1;
    var TempPerosonaArra = [];
    if (NewTempNumber >= 1) {
      var TempAllArray = this.state._items;
      var TempComplete = TempAllArray.filter(function (TempAllArray) {
        return TempAllArray["key"] == NewTempNumber;
      });

      var AuthorIdInArray = TempComplete[0]["AuthoName"];
      var ManagerIdInArray = TempComplete[0]["ManagerName"];
      var TeammembersIdsArray = TempComplete[0]["TeamMemebers"];
      var FinalName = "";


      if (isNaN(AuthorIdInArray) == false) {
        var UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == AuthorIdInArray;
        });
        TempComplete[0]["AuthoName"] = UserArray1[0]["Name"];
        FinalName = UserArray1[0]["Name"];
      } else {
        FinalName = TempComplete[0]["AuthoName"];
      }

      var NewObject1 = {
        imageInitials: '',
        text: FinalName,
        secondaryText: 'Software Engineer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      }
      TempPerosonaArra.push(NewObject1);



      if (isNaN(ManagerIdInArray) == false) {
        UserArray1 = this.state.UserArray;
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == ManagerIdInArray;
        });
        TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];
        FinalName = UserArray1[0]["Name"];
      }
      else {
        FinalName = TempComplete[0]["ManagerName"];
      }
      var NewObject2 = {
        imageInitials: '',
        text: FinalName,
        secondaryText: 'Software Engineer',
        tertiaryText: 'In a meeting',
        optionalText: 'Available at 4:00pm'
      }
      TempPerosonaArra.push(NewObject2);

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
        } else {
          TeamMemebrNameString.push(TeammembersIdsArray[y]);
        }
      }

      if (conversioncounter > 0) {
        TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;
      }

      this.setState(
        {
          CurrentIemCounter: NewTempNumber,
          SelectedItemId: TempComplete[0]["Id"],
          showPanel: true,
          SelectedItemArray: TempComplete,
          PersonaArray: TempPerosonaArra,
          PersonNameArray: TeamMemebrNameString,
        });

    }
  }

  public GetContcatCardManager() {
    //alert("manager");
  }

  public GetContactCardTeamMembers(items) {
    //alert(items);
  }

  private _onRenderExpandedCard = (item: any): JSX.Element => {
    var TempUserPropArray = this.state.UserArray;
    var ManagerName = this.state.PersonaArray[1]["text"];

    TempUserPropArray = TempUserPropArray.filter(function (TempUserPropArray) {
      return TempUserPropArray["Name"] == ManagerName;
    });

    return (
      <div className="hoverCardExample-expandedCard">

        <div className="ms-Grid-row">
          <div className={styles.PaddingSpace}>
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4">
              Organization
              </div>
            <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
              {TempUserPropArray[0]["Organization"]}
            </div>
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className={styles.PaddingSpace}>
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4">
              Tel
              </div>
            <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
              {TempUserPropArray[0]["Tel"]}
            </div>
          </div>
        </div>


        <div className="ms-Grid-row">
          <div className={styles.PaddingSpace}>
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4">
              Emails
              </div>
            <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
              {TempUserPropArray[0]["Email"]}
            </div>
          </div>
        </div>


        <div className="ms-Grid-row">
          <div className={styles.PaddingSpace}>
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4">
              Location
              </div>
            <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg8">
              {TempUserPropArray[0]["Location"]}
            </div>
          </div>
        </div>

      </div>
    );
  };


  private _onRenderCompactCard = (item: any): JSX.Element => {
    var TempUserPropArray = this.state.UserArray;
    var ManagerName = this.state.PersonaArray[1]["text"];

    TempUserPropArray = TempUserPropArray.filter(function (TempUserPropArray) {
      return TempUserPropArray["Name"] == ManagerName;
    });

    return (
      <div className="hoverCardExample-compactCard">
        <div className={styles.PaddingSpace}>
          <div className="ms-Grid" dir="ltr">

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg12">
                <h3>{TempUserPropArray[0]["Name"]}</h3>
              </div>
            </div>

            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg12">
                <Persona
                  {...this.state.PersonaArray[1]}
                  size={PersonaSize.size24}
                  presence={PersonaPresence.online}
                  hidePersonaDetails={false}
                />
              </div>
            </div>






          </div>
          <a target="_blank" >
          </a>
        </div>
      </div>
    );
  };

  public render(): React.ReactElement<IWPprojectdetailsProps> {

    const expandingCardProps: IExpandingCardProps = {
      onRenderCompactCard: this._onRenderCompactCard,
      onRenderExpandedCard: this._onRenderExpandedCard,
      renderData: this.state.PersonaArray[1],
      directionalHint: DirectionalHint.rightCenter,
      directionalHintFixed: true,
      gapSpace: 16,
      mode: ExpandingCardMode.expanded,
      expandedCardHeight: 200,
      compactCardHeight: 100,

    };


    if (this.state.PersonNameArray.length > 0) {
      var tmp = this.state.PersonNameArray;
      var PersonaPerSonsArray = tmp.map(function (item, i) {
        return (
          <div>
            <div onClick={this.GetContactCardTeamMembers(item)}>
              <Persona
                imageInitials=''
                text={item}
                size={PersonaSize.size24}
                presence={PersonaPresence.online}
                hidePersonaDetails={false}
              />
            </div>
          </div>
        );
      }, this);
    }

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
          <span>
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
                    <div className="ms-Grid-col ms-u-sm6 block">
                      {//this.state.SelectedItemArray[0]["AuthoName"]
                      }
                      <Persona
                        {...this.state.PersonaArray[0]}
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
                    <div className="ms-Grid-col ms-u-sm6 block" onClick={this.GetContcatCardManager.bind(this)}>
                      <HoverCard id="myID1" instantOpenOnClick={true} expandingCardProps={expandingCardProps}
                      >
                        <div className="HoverCard-item">
                          <Persona
                            {...this.state.PersonaArray[1]}
                            size={PersonaSize.size24}
                            presence={PersonaPresence.online}
                            hidePersonaDetails={false}
                          />
                        </div>
                      </HoverCard>
                    </div>

                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Team Members</div>
                    <div className="ms-Grid-col ms-u-sm6 block">
                      <div>
                        {PersonaPerSonsArray}
                      </div>
                    </div>
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
                    <div className="ms-Grid-col ms-u-sm6 block">Client</div>
                    <div className="ms-Grid-col ms-u-sm6 block">client@onmicrosoft.com</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Client Contact</div>
                    <div className="ms-Grid-col ms-u-sm6 block">client@onmicrosoft.com</div>
                  </div>
                </div>


                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project Start Date</div>
                    <div className="ms-Grid-col ms-u-sm6 block">client@onmicrosoft.com</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project End Date</div>
                    <div className="ms-Grid-col ms-u-sm6 block" >
                    </div>


                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">OpenAir Project Id</div>
                    <div className="ms-Grid-col ms-u-sm6 block">05151</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project Code</div>
                    <div className="ms-Grid-col ms-u-sm6 block">05151</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Planned Hours</div>
                    <div className="ms-Grid-col ms-u-sm6 block">15</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Planned Days</div>
                    <div className="ms-Grid-col ms-u-sm6 block">15</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Reporting Requirements</div>
                    <div className="ms-Grid-col ms-u-sm6 block">Reporting Requirements</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Project Deliverables</div>
                    <div className="ms-Grid-col ms-u-sm6 block">Project Deliverables</div>
                  </div>
                </div>

                <div className={styles.PaddingSpace} >
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-u-sm6 block">Useful Links</div>
                    <div className="ms-Grid-col ms-u-sm6 block">Useful Links</div>
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

  /********************** Dummy Data  */
  public itemdetailsDummy(id): void {
    
    var tmpString = id;
    var tmpid = parseInt(tmpString);
    var TempAllArray = this.state._items;
    var TempComplete = [];

    var TempPerosonaArra = [];

    TempComplete = TempAllArray.filter(function (TempAllArray) {
      return TempAllArray["id"] == tmpid;
    });


    var AuthorIdInArray = TempComplete[0]["AuthoName"];
    var ManagerIdInArray = TempComplete[0]["ManagerName"];
    var TeammembersIdsArray = TempComplete[0]["TeamMemebers"];

    var FinalName = "";
    var UserArray1 = this.state.UserArray;
    if (isNaN(AuthorIdInArray) == false) {
      UserArray1 = UserArray1.filter(function (UserArray1) {
        return UserArray1["Id"] == AuthorIdInArray;
      });
      TempComplete[0]["AuthoName"] = UserArray1[0]["Name"];
      FinalName = UserArray1[0]["Name"];
    } else {
      FinalName = TempComplete[0]["AuthoName"];
    }

    var NewObject1 = {
      imageInitials: '',
      text: FinalName,
      secondaryText: 'Software Engineer',
      tertiaryText: 'In a meeting',
      optionalText: 'Available at 4:00pm'
    }
    TempPerosonaArra.push(NewObject1);



    UserArray1 = this.state.UserArray;
    if (isNaN(ManagerIdInArray) == false) {
      UserArray1 = UserArray1.filter(function (UserArray1) {
        return UserArray1["Id"] == ManagerIdInArray;
      });
      TempComplete[0]["ManagerName"] = UserArray1[0]["Name"];
      FinalName = UserArray1[0]["Name"];
    } else {
      FinalName = TempComplete[0]["ManagerName"];
    }

    var NewObject2 = {
      imageInitials: '',
      text: FinalName,
      secondaryText: 'Software Engineer',
      tertiaryText: 'In a meeting',
      optionalText: 'Available at 4:00pm'
    }
    TempPerosonaArra.push(NewObject2);

    var Arobject = [];

    var TeamMemebrNameString = [];

    for (var y = 0; y < TeammembersIdsArray.length; y++) {

      var UserArray1 = this.state.UserArray;
      if (isNaN(TeammembersIdsArray[y]) == false) {
        UserArray1 = UserArray1.filter(function (UserArray1) {
          return UserArray1["Id"] == TeammembersIdsArray[y];
        });
        TeamMemebrNameString.push(UserArray1[0]["Name"]);
      } else {
        TeamMemebrNameString.push(TeammembersIdsArray[y]);
      }
    }
    TempComplete[0]["TeamMemebers"] = TeamMemebrNameString;
    TempPerosonaArra.push(Arobject);
    this.setState(
      {
        SelectedItemId: parseInt(tmpString),
        showPanel: true,
        SelectedItemArray: TempComplete,
        PersonaArray: TempPerosonaArra,
        PersonNameArray: TeamMemebrNameString,
        CurrentIemCounter: TempComplete[0]["key"],
      }
    )
  }

  /***********Dummy Data End********************/


}
