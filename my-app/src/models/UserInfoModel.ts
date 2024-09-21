/**
 * Defines the User Infor
 */
export interface UserInfo {
    // The user fullname
    name: string;

    // The user phone
    phone?: number | undefined;

    // The user mail
    mail?: string;

    //The user image
    photo: Blob;

}
