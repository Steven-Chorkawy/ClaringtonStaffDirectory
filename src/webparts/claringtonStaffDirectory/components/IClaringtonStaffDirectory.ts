import { IUser } from "../interface/IUser";

export interface IClaringtonStaffDirectoryProps {
    description: string;
    users: IUser[];
    clientMode: any;
    context: any;
}


//#region StaffGrid Component Interface.
export interface IStaffGridState {
    loadingUsers: boolean;  // Show or hide loading icons.
    persona: any;           // A list of users that will be rendered and visible to the end user.
    allPersonas: any;       // A list of all users that were queried from AD.  This array will not be rendered for the end user to see.
    // Used by detailed list.
    groups: any;            // Display users in a grouped detailed list.
    // Used by detailed list.
    columns: any;           // Format the detailed list that is going to render all the user.
}
//#endregion