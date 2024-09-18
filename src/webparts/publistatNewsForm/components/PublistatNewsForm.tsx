import * as React from "react";
import styles from "./PublistatNewsForm.module.scss";
import { IPublistatNewsFormProps } from "./IPublistatNewsFormProps";
import { escape } from "@microsoft/sp-lodash-subset";
import * as moment from "moment";
import {
  DefaultButton,
  Dialog,
  Icon,
  ITextFieldStyles,
  PrimaryButton,
  TextField,
} from "office-ui-fabric-react";
import { IFieldInfo, Social, sp } from "@pnp/sp/presets/all";
import { graph, Items } from "sp-pnp-js";

require("../assets/css/style.css");
require("../assets/css/fabric.min.css");

export interface IPublistatNewsFormState {
  AllNews: any;
  MyNews: any;
  status: string;
  AllUsers: any;
  Title: string;
  Date: any;
  Source: string;
  Link: string;
  Description: string;
  Pubdate: any;
  NewsGroup: string;
  Category: string;
  startDate: any;
  FilterDialog: boolean;
  EditFilterDialog: boolean;
  AddForm: boolean;
  AddFormTag: any;
  MySavedNews: any;
  EditNewsTitle: any;
  EditNewsSource: any;
  EditNewsDate: any;
  CurrentNewsID: any;
  CurrentUser: any;
  CurrentUserTitle: any;
  groups: any;
  CurrentUserGorupTitle: boolean;
  OwnerGroups: any;
  NewsSiteOwnerGroups: boolean;
  AllSiteGroupOwner: any;
  OwnerGroupMemeber: boolean;
}

const FilterDialogContentProps = {
  title: "Add News",
};

const EditFilterDialogContentProps = {
  title: "Update News",
};

// let columns = [
//   {
//     key: "Title",
//     name: "Title",
//     fieldName: "Title",
//     minWidth: 50,
//     maxWidth: 350,
//     isResizable: true,
//     onRender: (item) => {
//       return (
//         <a
//           href={item.Link}
//           style={{ textDecoration: "none", color: "#006eb5" }}
//         >
//           <div>
//             <span>
//               {item.Source}: {item.Title}
//             </span>
//           </div>
//         </a>
//       );
//     },
//   },
//   {
//     key: "Source",
//     name: "Source",
//     fieldName: "Source",
//     minWidth: 50,
//     maxWidth: 120,
//     isResizable: true,
//   },
//   {
//     key: "Pubdate",
//     name: "Publish Date",
//     fieldName: "Pubdate",
//     minWidth: 50,
//     maxWidth: 100,
//     isResizable: true,
//     onRender: (item) => {
//       return <span>{moment(new Date(item.Pubdate)).format("Do MMM")}</span>;
//     },
//   },
// ];

export default class PublistatNewsForm extends React.Component<
  IPublistatNewsFormProps,
  IPublistatNewsFormState
> {
  constructor(props: IPublistatNewsFormProps, state: IPublistatNewsFormState) {
    super(props);

    this.state = {
      AllUsers: "",
      AllNews: [],
      MyNews: [],
      Title: "",
      Date: "",
      Source: "",
      Link: "",
      Description: "",
      Pubdate: "",
      NewsGroup: "",
      Category: "",
      startDate: "",
      status: "",
      FilterDialog: true,
      EditFilterDialog: true,
      AddForm: true,
      AddFormTag: "",
      MySavedNews: [],
      EditNewsTitle: "",
      EditNewsSource: "",
      EditNewsDate: "",
      CurrentNewsID: "",
      CurrentUser: "",
      CurrentUserTitle: "",
      groups: "",
      CurrentUserGorupTitle: true,
      OwnerGroups: "",
      NewsSiteOwnerGroups: true,
      AllSiteGroupOwner: "",
      OwnerGroupMemeber: true,
    };
  }

  public render(): React.ReactElement<IPublistatNewsFormProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <div className="publistatNewsForm">
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-colms-sm12 ms-md12 ms-lg8 ms-xl6">
              <div className="d-flex-header">
                <h3 className="Title">News</h3>
                <div className="News-Add">
                  <PrimaryButton
                              text="Add"
                              onClick={() => this.setState({ FilterDialog: false })}
                      />
                </div>
            </div>
            </div> 

          <div className="ms-Grid-row flex-wrap-m">
            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg8 ms-xl8">
             
                {this.state.AllNews.length > 0 &&
                  this.state.AllNews.map((item) => {
                    return (
                      <>
                        <div className="ms-Grid-row News-Form">
                          <p className="ms-Grid NewsForm-date">
                            {moment(new Date(item.Date)).format("Do MM")}{" "}
                            <span>- {item.Category}</span>{" "}
                          </p>

                          <a
                            className="NewsLink"
                            href={item.Link}
                            data-interception="off"
                            target="_blank"
                          >
                            <h3 className="NewsForm-Title">
                              <span>{item.Source}</span> :{item.Title}
                            </h3>
                          </a>
                          {this.state.CurrentUserGorupTitle == true ? (
                            <>
                              <div className="ms-Grid-col NewsForm">
                                <div className="ms-Grid-col News-Edit">
                                  <PrimaryButton
                                    type="Edit"
                                    text="Edit"
                                    onClick={() =>
                                      this.setState(
                                        {
                                          EditFilterDialog: false,
                                          CurrentNewsID: item.ID,
                                        },
                                        () => this.GetNewsEditItem(item.ID)
                                      )
                                    }
                                  />
                                </div>
                                <div className="ms-Grid-col News-Delete">
                                  <DefaultButton
                                    type="Cancel"
                                    text="Delete"
                                    onClick={() => this.DeleteNews(item.ID)}
                                  />
                                </div>
                              </div>
                            </>
                          ) : (
                            <></>
                          )}
                        </div>
                      </>
                    );
                  })}
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md7 ms-lg5 ms-xl4">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg10">  
                    {/* <h3 className="areatitle">SiteOwnerGroups</h3> */}
                  </div>
                  {this.state.OwnerGroups.length > 0 &&
                    this.state.OwnerGroups.map((item) => {
                      return (
                        <>
                          <div className="Owner-Groups">
                            <div>{item.Title}</div>
                          </div>
                          {item.member.map((member) => {
                            return (
                              <>
                              <div className="Owner-Details">
                                <div className="Owner-img">
                                  <img src={ this.props.context.pageContext.web.absoluteUrl + `/_layouts/15/userphoto.aspx?UserName=${member.Title}&size=L`} draggable="false"/>
                                </div>
                                <div className="Owner-Member">
                                  <p>{member.Title}</p>
                                </div>
                              </div>
                              </>
                            );
                          })}
                        </>
                      );
                    })}
                </div>
              </div>
            </div>
          </div>
          <Dialog
            hidden={this.state.FilterDialog}
            onDismiss={() =>
              this.setState({
                FilterDialog: true,
                Title: "",
                Source: "",
                startDate: "",
              })
            }
            dialogContentProps={FilterDialogContentProps}
            minWidth={450}
          >
            <div>
              <div className="ms-Grid-row ms-md">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <TextField
                    label="Title"
                    name="Title"
                    type="text"
                    required={true}
                    onChange={(value) =>
                      this.setState({ Title: value.target["value"] })
                    }
                    value={this.state.Title}
                  />

                  <TextField
                    label="Source"
                    type="text"
                    name="Source"
                    onChange={(value) =>
                      this.setState({ Source: value.target["value"] })
                    }
                    value={this.state.Source}
                  />

                  <TextField
                    label="Date"
                    name="PubDate"
                    type="Date"
                    onChange={(value) =>
                      this.setState({ Date: value.target["value"] })
                    }
                    value={this.state.Date}
                  />
                </div>
              </div>
              <div className="ms-Grid-col Add-News">
                <div className="ms-Grid-col Submit-News">
                  <PrimaryButton
                    type="Submit"
                    text="Submit"
                    onClick={() => this.AddNews()}
                  />
                </div>
                <div className="ms-Grid-col Cancel-Add-News">
                  <DefaultButton
                    type="Cancel"
                    text="Cancel"
                    onClick={() => this.setState({ FilterDialog: true })}
                  />
                </div>
              </div>
            </div>
          </Dialog>

          <Dialog
            hidden={this.state.EditFilterDialog}
            onDismiss={() =>
              this.setState({
                EditFilterDialog: true,
                Title: "",
                Source: "",
                startDate: "",
              })
            }
            dialogContentProps={EditFilterDialogContentProps}
            minWidth={450}
          >
            <div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <TextField
                    label="Title"
                    name="Title"
                    type="text"
                    required={true}
                    onChange={(value) =>
                      this.setState({
                        EditNewsTitle: value.target["value"],
                      })
                    }
                    value={this.state.EditNewsTitle}
                  />

                  <TextField
                    label="Source"
                    type="text"
                    name="Source"
                    required={true}
                    onChange={(value) =>
                      this.setState({
                        EditNewsSource: value.target["value"],
                      })
                    }
                    value={this.state.EditNewsSource}
                  />

                  <TextField
                    label="Date"
                    name="PubDate"
                    type="Date"
                    required={true}
                    onChange={(value) =>
                      this.setState({
                        EditNewsDate: value.target["value"],
                      })
                    }
                    value={this.state.EditNewsDate}
                  />
                </div>
              </div>
              <div className="ms-Grid-col Edit-News">
                <div className="ms-Grid-col Update-News">
                  <PrimaryButton
                    type="Update"
                    text="Update"
                    onClick={() => this.UpdateNews(this.state.CurrentNewsID)}
                  />
                </div>
                <div className="ms-Grid-col Cancel-Update-News">
                  <DefaultButton
                    type="Cancel"
                    text="Cancel"
                    onClick={() => this.setState({ EditFilterDialog: true })}
                  />
                </div>
              </div>
            </div>
          </Dialog>
        </div>
      </div>
    );
  }

  public async componentDidMount() {
    await this.GetSiteOwnerGroup();
    this.GetNews();
    this.GetCurrentUserGroup();
  }

  public async GetCurrentUserGroup() {
    const site = await sp.web.siteGroups();
    let groups = await sp.web.currentUser.groups();

    console.log(site);
    console.log(groups);
    groups.forEach((items) => {
      if (items.Title == "PublistatNews Visitors") {
        this.setState({ CurrentUserGorupTitle: false });
      }
    });
    console.log(this.state.CurrentUserGorupTitle);
  }

  // public async GetSiteOwnerGroup() {
  //   const owenergroups= await sp.web.siteGroups();

  //   console.log(owenergroups);

  //   const ownergroupMemeber = await owenergroups.map(async (item) => {
  //     const member = await sp.web.siteGroups.getById(item.Id).users();
  //     return {
  //       ...item,
  //       member
  //     };
  //   })

  //   console.log(ownergroupMemeber);
  //   this.setState({ OwnerGroups : ownergroupMemeber});
  //   console.log(this.state.OwnerGroups);
  // }
  public async GetSiteOwnerGroup() {
    try {
      const owenergroups = await sp.web.siteGroups();

      console.log(owenergroups);

      const ownergroupMembers = await Promise.all(
        owenergroups.map(async (item) => {
          const member = await sp.web.siteGroups.getById(item.Id).users();
          return {
            ...item,
            member,
        
          };
        })
      );

      console.log(ownergroupMembers);
      this.setState({ OwnerGroups: ownergroupMembers });
      console.log(this.state.OwnerGroups);
    } catch (error) {
      console.error("Error fetching owner groups:", error);
    }
  }

  // owenergroupID.forEach((items) => {
  //   if(items.Title == " PublistatNews Owners") {
  //     this.setState({ AllSiteGroupOwner : "" })
  //   }
  // });
  // owenergroupmember.forEach((item) => {
  //   if(item.Title == "PublistatNews Members") {
  //     this.setState({ OwnerGroupMemeber : false })
  //   }
  // });
  // console.log(this.state.NewsSiteOwnerGroups);

  // let owenergroups = await sp.web.associatedOwnerGroup();
  // console.log(owenergroups);

  // const test = owenergroups
  // test.forEach((items) => {
  //   if(items.Title == "PublistatNews Visitors") {
  //     this.setState({ OwnerGroups : false})
  //   }
  // });

  public async AddNews() {
    if (
      this.state.Title.length == 0 ||
      this.state.Source.length == 0 ||
      this.state.Date.length == 0
    ) {
      alert("Please enter all the Details.");
    } else {
      const news = await sp.web.lists
        .getByTitle("News")
        .items.add({
          Title: this.state.Title,
          Source: this.state.Source,
          Date: this.state.Date,
        })
        .catch((error) => {
          console.log(error);
        });
      this.setState({ AllNews: "" });
      this.setState({ FilterDialog: true });
      this.GetNews();
    }
  }

  public async GetNewsEditItem(ID) {
    let EditNewsItem = this.state.AllNews.filter((item) => {
      if (item.ID == ID) {
        return item;
      }
    });
    console.log(EditNewsItem);
    this.setState({
      EditNewsTitle: EditNewsItem[0].Title,
      EditNewsSource: EditNewsItem[0].Source,
      EditNewsDate: EditNewsItem[0].Date,
    });
  }

  public async UpdateNews(CurrentNewsID) {
    const updatenews = await sp.web.lists
      .getByTitle("News")
      .items.getById(CurrentNewsID)
      .update({
        Title: this.state.EditNewsTitle,
        Source: this.state.EditNewsSource,
        Date: this.state.EditNewsDate,
      })
      .catch((error) => {
        console.log(error);
      });

    this.setState({ EditFilterDialog: true });
    this.GetNews();
  }

  public async DeleteNews(ID) {
    await sp.web.lists.getByTitle("News").items.getById(ID).delete();
    this.GetNews();
  }

  public async GetNews() {
    let items = [];
    let position = 0;
    const pageSize = 2000;
    let AllData = [];

    try {
      while (true) {
        const response = await sp.web.lists
          .getByTitle("News")
          .items.select(
            // "ID",
            "Title",
            "Link",
            "Pubdate",
            "Description",
            "Date",
            "Source",
            "Newsgroup",
            "Category",
            "Newsguid"
          )
          .orderBy("Date", false)
          .top(pageSize)
          .skip(position)
          .get();
        if (response.length === 0) {
          break;
        }
        items = items.concat(response);
        position += pageSize;
      }
      console.log(`Total items retrieved: ${items.length}`);
      if (items.length > 0) {
        items.forEach((item, i) => {
          AllData.push({
            ID: item.Id ? item.Id : "",
            Title: item.Title ? item.Title : "",
            Link: item.Link ? item.Link : "",
            Pubdate: item.Pubdate
              ? new Date(
                  new Date(item.Date).setHours(
                    new Date(item.Pubdate).getHours() + 2
                  )
                )
                  .toISOString()
                  .split("T")[0]
              : "",
            Description: item.Description ? item.Description : "",
            Date: item.Date
              ? new Date(
                  new Date(item.Date).setHours(
                    new Date(item.Date).getHours() + 2
                  )
                )
                  .toISOString()
                  .split("T")[0]
              : "",
            Source: item.Source ? item.Source : "",
            Newsgroup: item.Newsgroup ? item.Newsgroup : "",
            Category: item.Category ? item.Category : "",
          });
        });
        this.setState({ AllNews: AllData });
        console.log(this.state.AllNews);
      }
    } catch (error) {
      console.error(error);
    }
  }
}

// // <label>Update News</label>
// <label>Title</label>
// <label>Source</label>
// <label>Date</label>
// {/* <button>Edit</button> */}

// public updateNews() {
//   try{
//     const news = sp.web.lists.getByTitle("News")
//   }
// } catch (error) {
//   console.error(error);
// }
//    {/*----Delete News---- */}
//    {
//     this.state.AllNews.length > 0 &&
//     this.state.AllNews.map((item) => {
//       return (
//       <>
//         <div className='News-Form'>
//             <p className='NewsForm-Id'>
//               <span>{item.ID}</span>
//             </p>
//         </div>

//       </>
//     );
//   })
// }

//   <Dialog>
//         <TextField label='ID'
//               type="number"
//               name="ID"
//               onChange={(value) => this.state.AllNews()}
//           />
//           <div>
//             <PrimaryButton type="Delete" text="Trash" onClick={() => this.setState({ FilterDialog : true })}/>
//           </div>
//         </Dialog>

// // const groupName = "PublistatNews";
// const owenergroupName = await sp.web.siteGroups.getByName(groupName).users();
// console.log(owenergroupName);

// const groupName = "";
// const owenergroupName = await sp.web.siteGroups.getByName("PublistatNews").users();
