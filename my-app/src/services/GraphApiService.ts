import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { IGraphApiService } from "./IGraphApiService";
import { UserInfo } from "../models";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AppointmentModel, UTCDateRange } from "../models/AppointmentModel";
import { Event } from "../models/Event";
export class UserProfileService implements IGraphApiService {
  public static readonly serviceKey: ServiceKey<IGraphApiService> =
    ServiceKey.create<IGraphApiService>(
      "PnP:Teams:UserProfileService",
      UserProfileService
    );
  private _pageContext: PageContext | undefined;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
    });
  }

  /**
   * Get user info from GraphApi
   */
  public async getUserInfoByGraphApi(props: WebPartContext): Promise<UserInfo> {
    let result: UserInfo = {
      name: "",
      phone: undefined,
      mail: "",
      photo: new Blob(),
    };

    try {
      const client = await props.msGraphClientFactory.getClient("3");
      // user information API
      const userResponse = await client
        .api("/me")
        .version("v1.0")
        .select("displayName,businessPhones,mail")
        .get();

      result = {
        ...result,
        name: userResponse.displayName,
        phone: userResponse.businessPhones[0],
        mail: userResponse.mail,
      };
      // user photo API
      const photoResponse = await client
        .api("/me/photo/$value")
        .version("v1.0")
        .get();

      result.photo = photoResponse;
    } catch (err) {
      console.error(err);
    }

    return result;
  }

  public async getCalendarEvent(
    props: WebPartContext,
    dateRange: UTCDateRange
  ): Promise<AppointmentModel[]> {
    try {
      const client = await props.msGraphClientFactory.getClient("3");
      const response = await client
        .api(
          `/me/calendarview?startdatetime=${dateRange.start}&enddatetime=${dateRange.end}`
        )
        .expand(
          "extensions($filter=id eq 'Microsoft.OutlookServices.OpenTypeExtension.Com.Contoso.Referral')"
        )
        .get();
      const eventsData = response.value as AppointmentModel[];
      return eventsData;
    } catch (err) {
      console.error(err);
      return [];
    }
  }

  public async createCalendarEvent(
    props: WebPartContext,
    event: Event
  ): Promise<AppointmentModel[]> {
    try {
      const client = await props.msGraphClientFactory.getClient("3");
      const response = await client.api("/me/events").post(event);
      const eventsData = response.value as AppointmentModel[];
      return eventsData;
    } catch (err) {
      console.error(err);
      return [];
    }
  }
  public async updateCalendarEvent(
    props: WebPartContext,
    event: Event,
    id: string
  ): Promise<AppointmentModel[]> {
    try {
      const client = await props.msGraphClientFactory.getClient("3");
      const response = await client.api(`/me/events/${id}`).patch(event);
      const extensionResponse = await client
        .api(
          `/me/events/${id}/extensions/Microsoft.OutlookServices.OpenTypeExtension.Com.Contoso.Referral`
        )
        .patch(event.extensions[0]);
      const eventsData = response.value as AppointmentModel[];
      return eventsData;
    } catch (err) {
      console.error(err);
      return [];
    }
  }

  public async cancelAppointment(
    props: WebPartContext,
    id: string
  ): Promise<boolean> {
    try {
      const client = await props.msGraphClientFactory.getClient("3");
      await client.api(`/me/events/${id}`).delete();
      return true;
    } catch (err) {
      console.error("Error deleting event: ", err);
      return false;
    }
  }
  

  public async getAllUsers(props: WebPartContext): Promise<UserInfo[]> {
    const users: UserInfo[] = [];

    try {
      const client = await props.msGraphClientFactory.getClient("3");
      
      const userResponse = await client
        .api("/users")
        .version("v1.0")
        .select("id,displayName,businessPhones,mail")
        .get();

      for (const user of userResponse.value) {
        const userInfo: UserInfo = {
          name: user.displayName,
          phone: user.businessPhones ? user.businessPhones[0] : undefined,
          mail: user.mail,
          photo: new Blob(),
        };

        try {
          const photoReponse = await client
            .api("/users/${user.id}/photo/$value")
            .version("v1.0")
            .get();
          userInfo.photo = photoReponse;
        } catch {
          console.warn(`Can not get photo by ${user.displayName}`);
        }
        users.push(userInfo);
      }
    } catch (err) {
      console.error(err);
    }

    return users;
  }
}
