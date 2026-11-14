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
  selectedUsers,
  onUsersSelected,
  disable = false,
  itemLimit,
}) => {
  const [peoplePickerUsers, setPeoplePickerUsers] = React.useState<ITag[]>([]);

  // Get user photo URL exactly like SharePoint PeoplePicker does
  const getUserPhotoUrl = (userId: string, email: string): string => {
    const baseUrl = context.pageContext.web.absoluteUrl;
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

  // Fetch user suggestions with photos - Updated to include employee ID search
  const fetchEmails = async (filterText: string): Promise<ITag[]> => {
    if (!filterText || filterText.length < 3) return [];

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

      const usersWithPhotos: ITag[] = response.value.map((user: any) => ({
        key: user.id,
        name: user.displayName,
        email: user.mail,
        jobTitle: user.jobTitle,
        employeeId: user.employeeId,
        mobilePhone: user.mobilePhone,
        photoUrl: getUserPhotoUrl(user.id, user.mail),
      }));

      return usersWithPhotos;
    } catch (err) {
      console.error("Error fetching suggestions:", err);
      return [];
    }
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

  // Updated compact selected item renderer
  const onRenderSelectedItem = (props: IPickerItemProps<ITag>): JSX.Element => {
    const item = props.item;
    return (
      <div
        className="inline-flex items-center bg-blue-50 border border-blue-200 rounded-md px-2 py-1 mr-1 mb-1 text-lg"
        style={{
          fontSize: "13px",
          height: "28px",
          width: "260px",
          lineHeight: "22px",
        }}
      >
        <Persona
          text={item.name}
          imageUrl={item.photoUrl}
          size={PersonaSize.size16}
          hidePersonaDetails
          showInitialsUntilImageLoads={true}
          styles={{
            root: { minWidth: "18px" },
          }}
        />
        <span
          className="ml-1 text-gray-800 font-medium max-w-40 truncate"
          title={`${item.name}${item.employeeId ? ` (${item.employeeId})` : ""}`}
        >
          {item.name}
        </span>
        {item.employeeId && (
          <span className="ml-1 text-gray-500 text-xs">
            ({item.employeeId})
          </span>
        )}
        <button
          type="button"
          onClick={() => props.onRemoveItem && props.onRemoveItem()}
          title="Remove"
          className="ml-2 text-red-500 hover:text-red-600 transition text-sm hover:scale-105 font-bold leading-none"
          style={{
            fontSize: "13px",
            width: "13px",
            height: "13px",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          ×
        </button>
      </div>
    );
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    const allowedKeys = [
      "Backspace",
      "Delete",
      "ArrowLeft",
      "ArrowRight",
      "Tab",
      " ",
    ];
    const isLetter = /^[a-zA-Z0-9]$/.test(e.key);
    if (!isLetter && allowedKeys.indexOf(e.key) === -1) {
      e.preventDefault();
    }
  };

  return (
    <div onKeyDown={handleKeyDown}>
      <TagPicker
        styles={{
          root: {
            color: "black",
            minHeight: "32px",
          },
          input: {
            color: "black",
            fontSize: "13px",
          },
          text: {
            color: "black",
            fontSize: "13px",
          },
          itemsWrapper: {
            color: "black",
          },
        }}
        className="text-black dark:text-white"
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
          searchingText: "Searching users...",
        }}
        itemLimit={itemLimit === 1 ? 1 : undefined}
        disabled={disable}
        inputProps={{
          placeholder: "Search by name, email, or employee ID...",
        }}
      />
    </div>
  );
};

export default TagPeoplePicker;
