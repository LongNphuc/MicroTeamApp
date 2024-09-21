import { UserInfo } from "../models";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AppointmentModel,UTCDateRange } from "../models/AppointmentModel";
import { Event} from "../models/Event";
/**
 * Defines the abstract interface for the UserProfile Service
 */
export interface IGraphApiService {
     /**
     * Get user info from GraphApi
     */
    getUserInfoByGraphApi: (props : WebPartContext) => Promise<UserInfo>;
    getCalendarEvent(props: WebPartContext,dateRange: UTCDateRange): Promise<AppointmentModel[]> ;
    createCalendarEvent(props: WebPartContext,event: Event): Promise<AppointmentModel[]> ;
    updateCalendarEvent(props: WebPartContext,event: Event,id : string): Promise<AppointmentModel[]> ;
    cancelAppointment(props: WebPartContext,id : string): Promise<boolean>;


    //user
    getAllUsers: (props: WebPartContext) => Promise<UserInfo[]>;
}