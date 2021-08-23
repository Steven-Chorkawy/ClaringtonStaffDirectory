import { IUser } from "../interface/IUser";

export interface IClaringtonStaffDirectoryProps {
    description: string;
    users: IUser[];
    clientMode: any;
    context: any;
}

export interface IClaringtonStaffDirectoryState {
    users: IUser[];
    persona: any;
    allPersonas: any;    
    searchString?: string;
}