namespace OfficeDevPnP.Core {
    public enum SiteLockState {
        Unlock,
        NoAccess
    }

    public enum TenantOperationMessage
    {
        None,
        CreatingSiteCollection,
        DeletingSiteCollection,
        RemovingDeletedSiteCollectionFromRecycleBin,
        SettingSiteLockState,
        SettingSiteProperties
    }
}
