//import "@pnp/polyfill-ie11";
//import "react-app-polyfill/ie11";

import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users/web";
import { IUserProfileService } from "./IUserProfileService";
export class UserProfileService implements IUserProfileService {
    constructor(context: any) {
        sp.setup({
            ie11: true,
            spfxContext: context
        });
    }

    // Get Current Logged In User Properties
    public GetCurrentUserProperties = async (): Promise<any> => {
        return await this._getCurrentUserProperties();
    }
    // Get User Properties By Login Name
    public GetUserPropertiesByLoginName = async (loginName: string): Promise<any> => {
        return await this._getUserPropertiesByLoginName(loginName);
    }
    // Get Specific User Property by Login Name
    public GetSpecificUserProfileProperty = async (loginName: string, propertyName: string): Promise<string> => {
        return await this._getSpecificUserProfileProperty(loginName, propertyName);
    }
    // Get All SharePoint Groups 
    public GetSiteSPGroups = async (): Promise<any> => {
        return await this._getSiteSharePointGroups();
    }

    // Get all Users from respective SharePoint Group
    public GetUsersFromSharePointGroup = async (SPGroupId: number): Promise<any[]> => {
        return this._getUsersFromSharePointGroup(SPGroupId);
    }

    public CheckUsersExistsInSharePointGroup= async (SPGroupId: number,UserId : number): Promise<any> => {
        return this._checkUsersExistsInSharePointGroup(SPGroupId,UserId);
    }

    // Get all SharePoint Groups for a User
    public GetAllUsersGroupsFromSharePoint = async (): Promise<any[]> => {
        return this._getAllGroupsIdFromSharepointUserId();
    }
    private _getCurrentUserProperties = async (): Promise<any> => {
        let _userDetails: any;
        try {

            _userDetails = await sp.profiles.myProperties.get();

        }
        catch (err) {
            console.log(err);
        }
        return _userDetails;
    }

    private _getUserPropertiesByLoginName = async (loginName: string): Promise<any> => {
        let _userDetails: any;
        let loginNameFormat = "i:0#.f|membership|" + loginName;
        try {
            _userDetails = await sp.profiles.getPropertiesFor(loginNameFormat);
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _userDetails;
    }

    private _getSpecificUserProfileProperty = async (loginName: string, propertyName: string): Promise<string> => {
        let _userDetails: string = null;
        let loginNameFormat = "i:0#.f|membership|" + loginName;
        try {
            _userDetails = await sp.profiles.getUserProfilePropertyFor(loginNameFormat, propertyName);
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _userDetails;
    }

    private _getSiteSharePointGroups = async (): Promise<any> => {
        let _spGroups: any;
        try {
            let grps = await sp.web.siteGroups.get();
            _spGroups = grps.filter(g => {
                return !/^SharingLinks./.test(g.LoginName);
            });
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _spGroups;
    }

    private _getUsersFromSharePointGroup = async (SPGroupId: number): Promise<any[]> => {
        let _user: any[];
        try {
            _user = await sp.web.siteGroups.getById(SPGroupId).users.get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _user;
    }

    private _checkUsersExistsInSharePointGroup = async (SPGroupId: number,UserId:number): Promise<any> => {
        let _user: any;
        try {
            let users = await sp.web.siteGroups.getById(SPGroupId).users.get();
            
            _user = users.filter(item => item.Id === UserId);
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _user;
    }

    private _getAllGroupsIdFromSharepointUserId = async (): Promise<any> => {
        let _groups: any;
        try {
            _groups = await sp.web.currentUser.groups.get();
        }
        catch (err) {
            console.log(err);
            return null;
        }
        return _groups;
    }

}