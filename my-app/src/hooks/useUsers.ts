import { useState, useEffect } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { UserInfo } from "../models";

export const useUsers = (context: WebPartContext) => {
  const [users, setUsers] = useState<UserInfo[]>([]);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        const client = await context.msGraphClientFactory.getClient("3");
        const userResponse = await client
          .api("/users")
          .version("v1.0")
          .select("id,displayName,businessPhones,mail")
          .get();

        const usersData = await Promise.all(
          userResponse.value.map(async (user: any) => {
            let photo: Blob | undefined;
            try {
              const photoResponse = await client
                .api(`/users/${user.id}/photo/$value`)
                .version("v1.0")
                .get();
              photo = photoResponse;
            } catch (photoError) {
              console.warn(`Can not get photo by ${user.displayName}`);
            }

            return {
              name: user.displayName,
              phone: user.businessPhones ? user.businessPhones[0] : undefined,
              mail: user.mail,
              photo: photo,
            };
          })
        );

        setUsers(usersData);
      } catch (err) {
        console.error(err);
      }
    };

    fetchUsers().catch((err) => console.error("Error fetching users", err));
  }, [context]);

  return { users };
};
