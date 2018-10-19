import { Environment } from "../model/proxy";
import {
    ExchangeService, Uri, FolderId, WellKnownFolderName,
    Mailbox, ExchangeVersion, CalendarFolder,
    FolderView, SearchFilter, FolderSchema, FindFoldersResults,
    ExtendedPropertyDefinition, MapiPropertyType,
    GetUserSettingsResponse, UserSettingName, PropertySet,
    BasePropertySet, EmailMessage, ItemView, ItemTraversal,
    LogicalOperator,
    IOutParam,
    FindItemsResults,
    Item
} from "ews-javascript-api";
import { applyCredentials, calculateStoreId, validateAutodiscoverRedirection } from "../proxy/helper";
import { AutodiscoverService as NtlmAutodiscoverService } from '../extensions/CustomAutodiscoverService';
import { Buffer } from "buffer";

export class CreateWunderbarLinkRequest {

    async execute(env: Environment, params: { email: string }, data: { targetMailboxAddress: string, targetMailboxFolderId: string }) {
        let service = new ExchangeService(ExchangeVersion.Exchange2010);
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
        let PidTagWlinkGroupHeaderID = new ExtendedPropertyDefinition(0x6842, MapiPropertyType.Binary);
        let PidTagWlinkSaveStamp = new ExtendedPropertyDefinition(0x6847, MapiPropertyType.Integer);
        let PidTagWlinkOrdinal = new ExtendedPropertyDefinition(0x684B, MapiPropertyType.Binary);

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

        // 2B9065ED369E5C51BBDC2F7F8578942D => Confluence Calendars, GuidUtility.Create(GuidUtility.DnsNamespace, "com.yasoon.confluencecalendar.calendarGroup")
        // B9F0060000000000C000000000000046 => Shared Calendars
        // B8F0060000000000C000000000000046 => Other Calendars
        // B7F0060000000000C000000000000046 => My Calendars
        let confluenceGroupHeaderId = "2B9065ED369E5C51BBDC2F7F8578942D";
        let knownGroups = [confluenceGroupHeaderId, "B9F0060000000000C000000000000046", "B8F0060000000000C000000000000046", "B7F0060000000000C000000000000046"]
        let targetHeaderGuid = "B8F0060000000000C000000000000046";
        let calendarFolderType = Buffer.from("0278060000000000C000000000000046", "hex");
        let headers: FindItemsResults<Item>;
        try {
            //Find correct group header    
            // => FolderTypes: https://msdn.microsoft.com/en-us/library/ee218711(v=exchg.80).aspx        
            let associatedGroupHeaderView = new ItemView(100);
            associatedGroupHeaderView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, PidTagWlinkGroupHeaderID);
            associatedGroupHeaderView.Traversal = ItemTraversal.Associated;
            let groupHeaderSearchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
            groupHeaderSearchFilter.Add(new SearchFilter.IsEqualTo(new ExtendedPropertyDefinition(0x6849, MapiPropertyType.Integer), 4));
            groupHeaderSearchFilter.Add(new SearchFilter.IsEqualTo(new ExtendedPropertyDefinition(0x684F, MapiPropertyType.Binary), calendarFolderType.toString('base64')));
            headers = await commonViewsFolder.FindItems(groupHeaderSearchFilter, associatedGroupHeaderView);

            let confluenceGroups = headers.Items.filter(i => {
                let propValue: IOutParam<any> = <any>{};
                if (i.ExtendedProperties.TryGetValue(PidTagWlinkGroupHeaderID, propValue)) {
                    if (propValue.outValue)
                        return new Buffer(propValue.outValue).toString('hex').toUpperCase() === confluenceGroupHeaderId;
                }

                return false;
            });

            if (confluenceGroups.length === 0) {
                //According to https://msdn.microsoft.com/en-us/library/ee217241(v=exchg.80).aspx ([MS-OXOCFG] 4.4.1)
                let newGroup = new EmailMessage(service);
                newGroup.ItemClass = "IPM.Microsoft.WunderBar.Link";
                newGroup.IsAssociated = true;
                newGroup.Subject = "Confluence Calendars";
                newGroup.SetExtendedProperty(PidTagWlinkGroupHeaderID, Buffer.from(confluenceGroupHeaderId, 'hex'));
                newGroup.SetExtendedProperty(PidTagWlinkSaveStamp, Math.floor(Math.random() * 10000000));
                newGroup.SetExtendedProperty(PidTagWlinkType, 4); //wblHeader
                newGroup.SetExtendedProperty(PidTagWlinkFlags, 0);
                newGroup.SetExtendedProperty(PidTagWlinkOrdinal, Buffer.from('80', 'hex'));
                newGroup.SetExtendedProperty(PidTagWlinkFolderType, calendarFolderType);
                newGroup.SetExtendedProperty(PidTagWlinkSection, 3);
                await newGroup.Save(commonViewsFolder.Id);
            }

            targetHeaderGuid = confluenceGroupHeaderId;
        }
        catch (e) {
            //Try falling back to known groups
            try {
                for (let i = 0; i < knownGroups.length; i++) {
                    let matchedGroups = headers.Items.filter(item => {
                        let propValue: IOutParam<any> = <any>{};
                        if (item.ExtendedProperties.TryGetValue(PidTagWlinkGroupHeaderID, propValue)) {
                            if (propValue.outValue)
                                return new Buffer(propValue.outValue).toString('hex').toUpperCase() === knownGroups[i];
                        }

                        return false;
                    });

                    if (matchedGroups.length > 0) {
                        targetHeaderGuid = knownGroups[i];
                        break;
                    }
                }
            }
            catch (e) { }
        }

        objWunderBarLink.ItemClass = "IPM.Microsoft.WunderBar.Link";
        //objWunderBarLink.SetExtendedProperty(PidTagWlinkAddressBookEID, Buffer.from(abTargetABEntryId, "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkAddressBookStoreEID, Buffer.from(ownStoreId, "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkCalendarColor, -1);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkFlags, 0);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkGroupName, "Shared Calendars");
        objWunderBarLink.SetExtendedProperty(PidTagWlinkFolderType, calendarFolderType);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkGroupClsid, Buffer.from(targetHeaderGuid, "hex"));
        objWunderBarLink.SetExtendedProperty(PidTagWlinkROGroupType, -1);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkSection, 3);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkType, 2);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkEntryId, sharedEntryId);
        objWunderBarLink.SetExtendedProperty(PidTagWlinkStoreEntryId, Buffer.from(targetStoreId, "hex"));
        objWunderBarLink.IsAssociated = true;

        await objWunderBarLink.Save(commonViewsFolder.Id);
    }
}