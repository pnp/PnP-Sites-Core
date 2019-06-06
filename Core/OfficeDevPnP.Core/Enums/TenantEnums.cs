namespace OfficeDevPnP.Core {
    public enum SiteLockState {
        Unlock,
        NoAccess,
        ReadOnly
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
