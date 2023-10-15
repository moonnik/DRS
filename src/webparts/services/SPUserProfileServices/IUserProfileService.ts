
export interface IUserProfileService {
    GetCurrentUserProperties: () => Promise<any>;
    GetUserPropertiesByLoginName: (loginName: string) => Promise<any>;
    GetSpecificUserProfileProperty: (loginName: string, propertyName: string) => Promise<string>;
    GetSiteSPGroups: () => Promise<any>;
    GetUsersFromSharePointGroup: (SPGroupId: number) => Promise<any[]>;
    CheckUsersExistsInSharePointGroup: (SPGroupId: number, UserId: number) => Promise<any[]>;
    GetAllUsersGroupsFromSharePoint: () => Promise<any[]>;
}