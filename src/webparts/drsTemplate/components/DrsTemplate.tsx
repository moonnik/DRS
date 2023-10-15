import * as React from 'react';
import styles from './DrsTemplate.module.scss';
import { IDrsTemplateProps } from './IDrsTemplateProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Route, Switch, HashRouter } from 'react-router-dom';
import { Create, Person } from "@material-ui/icons";
import { Dashboard as DashboardIcon } from '@material-ui/icons';
import Button from '@material-ui/core/Button';
import ButtonGroup from '@material-ui/core/ButtonGroup';
import LinearProgress from '@material-ui/core/LinearProgress';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Dashboard } from './Dashboard/Dashboard';
import { NewRequest } from './New Request/NewRequest';
import { IUserProfileService, UserProfileService } from "../../services/SPUserProfileServices";

export default class DrsTemplate extends React.Component<IDrsTemplateProps, any, {}> {
  private userProfileServices: IUserProfileService;

  private isAllRequestBtnDisabled = window.location.hash.toLowerCase().indexOf("#/requests/all") > -1 ? true : false;
  private isNewReqBtnDisabled = window.location.hash.toLowerCase() == "#/request/new" ? true : false;

  constructor(props: IDrsTemplateProps) {
    super(props);
    this.state = {
      currentUser: null,
      loading: true,
      showCreateButton: true,
      isValidUser: true,
      NewItem: {},
      IsRequestor: true

    };
    this.userProfileServices = new UserProfileService(this.props.context);
    // this.OnLoad();
  }

  private CreateNew(this) {

    this.isAllRequestBtnDisabled = (window.location.hash.toLowerCase() == "#/" || window.location.hash.toLowerCase().indexOf("#/requests/all") > -1) ? false : false;
    this.isNewReqBtnDisabled = true;
    this.setState({ showCreateButton: false });
    window.location.hash = '#/request/new';
  }
  private Allrequest() {

    this.isAllRequestBtnDisabled = true;
    this.isNewReqBtnDisabled = window.location.hash.toLowerCase() == "#/requests/new" ? false : false;
    this.setState({ showCreateButton: true });
    window.location.hash = '#/';
  }

  public render(): React.ReactElement<IDrsTemplateProps> {
    return (
      <div>
        <HashRouter>

          <div className={styles.drsTemplate}>
            <div className={styles.row}>
              <div>
                <div className={styles.header}>
                  Decision Record Sheet

                </div>
                <div className={styles.buttonTop}>
                  <ButtonGroup variant="contained" color="primary" aria-label="contained primary button group" >
                    <Button id="btnTopAllReq" startIcon={<DashboardIcon />} className={styles.ButtonMain}  onClick={this.Allrequest.bind(this)}>DRS REQUESTS</Button>
                    {/* <Button id="btnTopMyReq" startIcon={<Person />} className={styles.ButtonMain} disabled={this.isMyRequestBtnDisabled} onClick={this.Myrequest.bind(this)}>MY REQUESTS</Button> */}
                    <Button id="btnTopNewReq" startIcon={<Create />} className={styles.ButtonMain} onClick={this.CreateNew.bind(this)}>CREATE NEW REQUEST</Button>
                  </ButtonGroup>
                </div>
              </div>
            </div>
          </div>
          {(this.state.loading == true ?
            <div className={styles.container}>
              {this.state.isValidUser ?
                <Switch>
                  <Route path="/" exact={true} render={({ match, location }) => {
                    return <Dashboard
                      displayItemsCount={100}
                      DRSLibraryID={this.props.LibraryId}
                      listServices={this.props.listServices}
                      CurrentUser={this.props.context.pageContext.legacyPageContext.userId}
                      folderPath={this.props.folderPath}
                      Filter={this.props.Filter}
                      DRSApprovalsListId={this.props.DRSApprovalsListId}
                      context={this.props.context}
                    />
                  }
                  }
                  />

                  <Route path="/requests/my" exact={true} render={({ match, location }) => {
                    return < Dashboard
                      displayItemsCount={100}
                      DRSLibraryID={this.props.LibraryId}
                      listServices={this.props.listServices}
                      CurrentUser={this.props.context.pageContext.legacyPageContext.userId}
                      isMyRequests={true}
                      folderPath={this.props.folderPath}
                      DRSApprovalsListId={this.props.DRSApprovalsListId}
                      context={this.props.context}
                    />
                  }} />
                  <Route path="/requests/all" exact={true} render={({ match, location }) => {
                    return <Dashboard
                      displayItemsCount={100}
                      DRSLibraryID={this.props.LibraryId}
                      listServices={this.props.listServices}
                      CurrentUser={this.props.context.pageContext.legacyPageContext.userId}
                      folderPath={this.props.folderPath}
                      Filter={this.props.Filter}
                      DRSApprovalsListId={this.props.DRSApprovalsListId}
                      context={this.props.context}
                    />
                  }} />

                  <Route path="/request/new" exact={true}
                    render={({ match }) => {
                      return this.state.IsRequestor ? <NewRequest
                        DRSLibray={this.props.LibraryId}
                        FlowURL={this.props.FlowURL}
                        context={this.props.context}
                        listServices={this.props.listServices}
                        ProjectListId={this.props.ProjectListId}
                       // ProjectNumberListId={this.props.ProjectNumberListId}


                      /> : <Placeholder
                        iconName='WarningSolid'
                        iconText='Access Denied'
                        description='You do not have permission to create a new request. Please contact Site Administrator to request access.'>
                      </Placeholder>;
                    }}
                  />

                  {/* <Route
                    exact
                    path="/request/view/:Id"
                    render={({ match }) => {
                      return <ViewRequest
                        RequestListID={this.props.RequestList}
                        ID={match.params.Id}
                        context={this.props.context}
                        
                      />;
                    }}
                  />              

                  <Route
                    exact
                    path="/request/edit/:Id"
                    render={({ match }) => {
                      return this.state.IsRequestor ? <EditRequest
                        RequestListID={this.props.RequestList}
                        ID={match.params.Id}
                        ListServices={this.props.listServices}
                        context={this.props.context}                        
                        isAdmin={this.state.isAdminMember}
                      //CPDSPGroupId={this.props.CPDSPGroupId}                         
                      ></EditRequest> : <Placeholder iconName='WarningSolid'
                        iconText='Access Denied'
                        description='You do not authorized to edit the request.'
                      ></Placeholder>;
                    }}
                  />
*/}
                  <Route path="*">
                    <Placeholder
                      iconName='WarningSolid'
                      iconText='Invalid URL'
                      description='Please check the URL.'>
                    </Placeholder>
                  </Route>
                </Switch>
                :
                <div>
                  <div className={styles.row}>
                    <Placeholder iconName='WarningSolid'
                      iconText='Access Denied'
                      description='You do not have permission to access this page. Please contact Site Administrator to request access.'
                    />
                  </div>
                </div>}

            </div>
            : <div>
              <LinearProgress color="primary" />
            </div>)}

        </HashRouter>
      </div>
    );
  }
}
