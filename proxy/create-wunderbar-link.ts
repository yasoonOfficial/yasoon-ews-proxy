import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, FolderId, WellKnownFolderName,
    Mailbox, ExchangeVersion, CalendarFolder,
    FolderView, SearchFilter, FolderSchema, FindFoldersResults,
    ExtendedPropertyDefinition, MapiPropertyType,
    GetUserSettingsResponse, UserSettingName, PropertySet,
    BasePropertySet, EmailMessage
} from "ews-javascript-api";
import { applyCredentials, calculateStoreId, validateAutodiscoverRedirection } from "../proxy/helper";
import { AutodiscoverService as NtlmAutodiscoverService } from '../extensions/CustomAutodiscoverService';

export class CreateWunderbarLinkRequest {

    async execute(env: Environment, params: { email: string }, data: { targetMailboxAddress: string, targetMailboxFolderId: string }) {
        let service = new ExchangeService(ExchangeVersion.Exchange2013);
        service.Url = new Uri(env.ewsUrl);
        applyCredentials(service, env);

        let ownUserEmail = params.email;
        let targetMailboxEmail = data.targetMailboxAddress;
        let targetMailboxFolderId = data.targetMailboxFolderId;

        let rootFolder = new FolderId(WellKnownFolderName.Root, new Mailbox(ownUserEmail));
        let commonViewFolderView = new FolderView(1000);
        let commonViewSearchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Common Views");
        let ewsResult: FindFoldersResults = await service.FindFolders(rootFolder, commonViewSearchFilter, commonViewFolderView);
        let commonViewsFolder = ewsResult.Folders[0];

        //Constants
        let PidTagWlinkFolderType = new ExtendedPropertyDefinition(0x684F, MapiPropertyType.Binary);
        let PidTagWlinkGroupName = new ExtendedPropertyDefinition(0x6851, MapiPropertyType.String);
        let pidTagEntryId = new ExtendedPropertyDefinition(4095, MapiPropertyType.Binary);
        let PidTagWlinkType = new ExtendedPropertyDefinition(0x6849, MapiPropertyType.Integer);
        let PidTagWlinkFlags = new ExtendedPropertyDefinition(0x684A, MapiPropertyType.Integer);
        let PidTagWlinkSection = new ExtendedPropertyDefinition(0x6852, MapiPropertyType.Integer);
        let PidTagWlinkStoreEntryId = new ExtendedPropertyDefinition(0x684E, MapiPropertyType.Binary);
        let PidTagWlinkGroupClsid = new ExtendedPropertyDefinition(0x6850, MapiPropertyType.Binary);
        let PidTagWlinkEntryId = new ExtendedPropertyDefinition(0x684C, MapiPropertyType.Binary);
        let PidTagWlinkCalendarColor = new ExtendedPropertyDefinition(0x6853, MapiPropertyType.Integer);
        let PidTagWlinkROGroupType = new ExtendedPropertyDefinition(0x6892, MapiPropertyType.Integer);
        let PidTagWlinkAddressBookStoreEID = new ExtendedPropertyDefinition(0x6891, MapiPropertyType.Binary);

        //Configure Autodiscover Service
        let autodiscoverService = new NtlmAutodiscoverService();

        //@ts-ignore
        autodiscoverService.RedirectionUrlValidationCallback = validateAutodiscoverRedirection;
        applyCredentials(autodiscoverService, env);

        let ownUserSettings: GetUserSettingsResponse = await autodiscoverService.GetUserSettings(
            ownUserEmail,
            UserSettingName.InternalRpcClientServer,
            UserSettingName.UserDN
        );

        let targetMailBoxSettings: GetUserSettingsResponse = await autodiscoverService.GetUserSettings(
            targetMailboxEmail,
            UserSettingName.UserDN,
            UserSettingName.InternalRpcClientServer,
            UserSettingName.UserDisplayName
        );

        let ownStoreId = calculateStoreId(ownUserSettings.Settings[UserSettingName.UserDN], ownUserSettings.Settings[UserSettingName.InternalRpcClientServer]);
        let targetStoreId = calculateStoreId(targetMailBoxSettings.Settings[UserSettingName.UserDN], targetMailBoxSettings.Settings[UserSettingName.InternalRpcClientServer]);
        //let abTargetABEntryId = calculateAddressBookId(targetMailBoxSettings.Settings[UserSettingName.UserDN]);
        let sharedCalFolderId = new FolderId();

        if (targetMailboxFolderId === 'main') {
            sharedCalFolderId = new FolderId(WellKnownFolderName.Calendar, new Mailbox(targetMailboxEmail));
        } else {
            sharedCalFolderId.UniqueId = targetMailboxFolderId;
        }

        let sharedCal = await CalendarFolder.Bind(service, sharedCalFolderId, new PropertySet(BasePropertySet.FirstClassProperties, [pidTagEntryId]));
        let sharedEntryId = sharedCal.ExtendedProperties.GetOrAddExtendedProperty(pidTagEntryId).Value;
        let objWunderBarLink = new EmailMessage(service);

        if (targetMailboxFolderId === 'main') {
            objWunderBarLink.Subject = targetMailBoxSettings.Settings[UserSettingName.UserDisplayName];
        } else {
            objWunderBarLink.Subject = sharedCal.DisplayName;
        }

        objWunderBarLink.ItemClass = "IPM.Microsoft.WunderBar.Link";
        //objWunderBarLink.SetExtendedProperty(PidTagWlinkAddressBookEID, Buffer.from(abTargetABEntryId, "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkAddressBookStoreEID, Buffer.from(ownStoreId, "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkCalendarColor, -1);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkFlags, 0);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkGroupName, "Shared Calendars");
        objWunderBarLink.SetExtendedProperty(PidTagWlinkFolderType, Buffer.from("0278060000000000C000000000000046", "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkGroupClsid, Buffer.from("B9F0060000000000C000000000000046", "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkROGroupType, -1);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkSection, 3);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkType, 2);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkEntryId, sharedEntryId);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkStoreEntryId, Buffer.from(targetStoreId, "hex"));
        objWunderBarLink.IsAssociated = true;

        await objWunderBarLink.Save(commonViewsFolder.Id);
    }
}