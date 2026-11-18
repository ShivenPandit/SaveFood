import * as React from "react";
import {
  TagPicker,
  Persona,
  PersonaSize,
  IPickerItemProps,
  ITag as IFluentTag,
} from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IUser {
  Id: string;
  title: string;
  email: string;
  jobTitle?: string;
  employeeId?: string;
  mobilePhone?: string;
}

export interface ITag extends IFluentTag {
  key: string;
  name: string;
  email?: string;
  jobTitle?: string;
  employeeId?: string;
  mobilePhone?: string;
  photoUrl?: string;
}

export interface ITagPickerProps {
  context: WebPartContext;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  selectedUsers: IUser[];
  onUsersSelected: (users: IUser[]) => void;
  disable?: boolean;
  itemLimit?: number;
}

const TagPeoplePicker: React.FC<ITagPickerProps> = ({
  context,
  spHttpClient,
  siteUrl,
  selectedUsers,
  onUsersSelected,
  disable = false,
  itemLimit,
}) => {
  const [peoplePickerUsers, setPeoplePickerUsers] = React.useState<ITag[]>([]);

  // Get user photo URL exactly like SharePoint PeoplePicker does
  const getUserPhotoUrl = (userId: string, email: string): string => {
    const baseUrl = siteUrl || context.pageContext.web.absoluteUrl;
    return `${baseUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${email}`;
  };

  // Enrich or map selected users
  React.useEffect(() => {
    const enrichUsers = async (): Promise<void> => {
      if (!selectedUsers || selectedUsers.length === 0) {
        setPeoplePickerUsers([]);
        return;
      }

      const enrichedTags: ITag[] = [];

      for (const user of selectedUsers) {
        if (user.Id && user.email) {
          enrichedTags.push({
            key: user.Id,
            name: user.title,
            email: user.email,
            jobTitle: user.jobTitle,
            employeeId: user.employeeId,
            mobilePhone: user.mobilePhone,
            photoUrl: getUserPhotoUrl(user.Id, user.email),
          });
        } else {
          try {
            const client = await context.msGraphClientFactory.getClient("3");
            const res = await client
              .api(`/users/${user.email}`)
              .select("id,displayName,mail,jobTitle,employeeId,mobilePhone")
              .get();

            enrichedTags.push({
              key: res.id,
              name: res.displayName,
              email: res.mail,
              jobTitle: res.jobTitle,
              employeeId: res.employeeId,
              mobilePhone: res.mobilePhone,
              photoUrl: getUserPhotoUrl(res.id, res.mail),
            });
          } catch (error) {
            console.warn("User not found in Graph:", user.email);
          }
        }
      }

      setPeoplePickerUsers(enrichedTags);
    };

    void enrichUsers();
  }, [selectedUsers, context]);

  // SharePoint Native Client People Picker Fallback (works in all tenants without Graph consent)
  const fetchFromSharePoint = async (filterText: string): Promise<ITag[]> => {
    try {
      const targetSiteUrl = siteUrl || context.pageContext.web.absoluteUrl;
      const url = `${targetSiteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface/ClientPeoplePickerSearchUser`;
      const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
        },
        body: JSON.stringify({
          queryParams: {
            __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
            AllowEmailAddresses: true,
            AllowMultipleEntities: true,
            AllUrlZones: false,
            MaximumEntitySuggestions: 30,
            PrincipalSource: 15,
            PrincipalType: 1,
            QueryString: filterText,
          },
        }),
      });

      if (response.ok) {
        const json = await response.json();
        const rawResults = JSON.parse(json?.d?.ClientPeoplePickerSearchUser || "[]");
        return rawResults.map((u: any) => {
          const email = u.EntityData?.Email || u.Description || "";
          return {
            key: u.Key || email || u.DisplayText,
            name: u.DisplayText,
            email: email,
            jobTitle: u.EntityData?.Title || u.EntityData?.Department,
            employeeId: u.EntityData?.EmployeeID,
            mobilePhone: u.EntityData?.MobilePhone,
            photoUrl: getUserPhotoUrl(u.Key, email),
          };
        });
      }
    } catch (spErr) {
      console.warn("SharePoint People Search fallback failed:", spErr);
    }
    return [];
  };

  // Fetch user suggestions with photos - Supports Microsoft Graph with SharePoint fallback
  const fetchEmails = async (filterText: string): Promise<ITag[]> => {
    if (!filterText || filterText.trim().length < 2) return [];

    try {
      const client = await context.msGraphClientFactory.getClient("3");

      // Enhanced filter to include employee ID search
      const response = await client
        .api("/users")
        .filter(
          `startswith(displayName,'${filterText}') or startswith(mail,'${filterText}') or startswith(employeeId,'${filterText}')`
        )
        .select("id,displayName,mail,jobTitle,employeeId,mobilePhone")
        .top(50)
        .get();

      if (response && response.value && response.value.length > 0) {
        const usersWithPhotos: ITag[] = response.value.map((user: any) => ({
          key: user.id || user.mail,
          name: user.displayName,
          email: user.mail,
          jobTitle: user.jobTitle,
          employeeId: user.employeeId,
          mobilePhone: user.mobilePhone,
          photoUrl: getUserPhotoUrl(user.id, user.mail),
        }));

        return usersWithPhotos;
      }
    } catch (err) {
      // MS Graph error (e.g. unconsented permissions) - seamlessly fall back to SharePoint People Search
    }

    return await fetchFromSharePoint(filterText.trim());
  };

  // Render each suggestion
  const onRenderSuggestion = (item: ITag): JSX.Element => (
    <div className="flex items-center p-1 rounded-lg hover:bg-gray-100 transition">
      <Persona
        text={item.name}
        secondaryText={item.email}
        tertiaryText={
          item.employeeId ? `ID: ${item.employeeId}` : item.jobTitle
        }
        size={PersonaSize.size32}
        imageUrl={item.photoUrl}
        hidePersonaDetails={false}
        showInitialsUntilImageLoads={true}
      />
    </div>
  );

  // Compact selected user pill renderer
  const onRenderSelectedItem = (props: IPickerItemProps<ITag>): JSX.Element => {
    const item = props.item;
    return (
      <div
        className="inline-flex items-center bg-blue-50/90 border border-blue-200 rounded-lg px-2 py-0.5 mr-1.5 my-0.5 text-[12.5px] shadow-2xs select-none"
        style={{
          height: "26px",
          maxWidth: "210px",
          lineHeight: "18px",
        }}
      >
        <Persona
          text={item.name}
          imageUrl={item.photoUrl}
          size={PersonaSize.size24}
          hidePersonaDetails
          showInitialsUntilImageLoads={true}
          styles={{
            root: { minWidth: "24px" },
          }}
        />
        <span
          className="ml-1.5 text-slate-800 font-semibold max-w-[120px] truncate text-[12.5px]"
          title={`${item.name}${item.email ? ` (${item.email})` : ""}`}
        >
          {item.name}
        </span>
        {item.employeeId && (
          <span className="ml-1 text-slate-400 text-[10.5px]">
            ({item.employeeId})
          </span>
        )}
        <button
          type="button"
          onClick={() => props.onRemoveItem && props.onRemoveItem()}
          title="Remove"
          className="ml-1.5 text-slate-400 hover:text-rose-600 transition text-sm font-bold leading-none cursor-pointer"
        >
          ×
        </button>
      </div>
    );
  };

  return (
    <div className="w-full">
      <TagPicker
        styles={{
          root: {
            color: "#1e293b",
            minHeight: "42px",
            borderRadius: "10px",
            borderColor: "#cbd5e1",
            backgroundColor: "#ffffff",
          },
          input: {
            color: "#0f172a",
            fontSize: "14px",
            fontWeight: 500,
            padding: "5px 8px",
            minWidth: "140px",
          },
          text: {
            color: "#0f172a",
            fontSize: "14px",
            minHeight: "42px",
            maxHeight: "110px",
            overflowY: "auto",
            display: "flex",
            alignItems: "center",
            padding: "2px 8px",
            borderColor: "#cbd5e1",
            borderRadius: "10px",
          },
          itemsWrapper: {
            padding: "2px",
            display: "flex",
            flexWrap: "wrap",
            alignItems: "center",
          },
        }}
        className="text-slate-800"
        onResolveSuggestions={fetchEmails}
        selectedItems={peoplePickerUsers}
        onChange={(items?: IFluentTag[]) => {
          const updatedTags = (items || []) as ITag[];
          setPeoplePickerUsers(updatedTags);

          const userMap: { [key: string]: IUser } = {};
          updatedTags.forEach((item) => {
            userMap[item.key] = {
              Id: String(item.key),
              title: item.name,
              email: item.email || "",
              jobTitle: item.jobTitle,
              employeeId: item.employeeId,
              mobilePhone: item.mobilePhone,
            };
          });

          const userList: IUser[] = Object.keys(userMap).map(
            (key) => userMap[key]
          );
          onUsersSelected(userList);
        }}
        onRenderSuggestionsItem={(item: IFluentTag) =>
          onRenderSuggestion(item as ITag)
        }
        onRenderItem={(props: IPickerItemProps<IFluentTag>) =>
          onRenderSelectedItem(props as IPickerItemProps<ITag>)
        }
        pickerSuggestionsProps={{
          suggestionsHeaderText: "Suggested Users",
          noResultsFoundText: "No matching users found",
          searchingText: "Searching directory...",
        }}
        itemLimit={itemLimit === 1 ? 1 : undefined}
        disabled={disable}
        inputProps={{
          placeholder: peoplePickerUsers.length === 0 ? "Search by name or email..." : "",
        }}
      />
    </div>
  );
};

export default TagPeoplePicker;
