namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Describe list provisioning stage.
    /// </summary>
    /// <remarks>Becase lookup fields require to have the target list available, and dependent lookups fields require having the primary lookup field
    /// available, theses stages allows running three times the provider objects, each time focusing on specific kind of artefact to create </remarks>
    public enum FieldStage
    {
        /// <summary>
        /// The list itself and fields that aren't lookup fields are provisioned
        /// </summary>
        ListAndStandardFields,
        /// <summary>
        /// Focus on lookup fields. This assumes target lists are yet available
        /// </summary>
        LookupFields,
        /// <summary>
        /// Focus on dependent lookup fields. This assumes primary lookup fields are available
        /// </summary>
        DependentLookupFields
    }
}