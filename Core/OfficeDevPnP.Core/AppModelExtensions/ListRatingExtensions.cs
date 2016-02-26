using System;
using System.Linq;
using OfficeDevPnP.Core;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Enables: Ratings / Likes functionality on list in publishing web.
    /// </summary>
    public static partial class ListRatingExtensions
    {
        /// TODO: Replace Logging throughout 

        #region Rating Field

        private static readonly Guid RatingsFieldGuid_AverageRating = new Guid("5a14d1ab-1513-48c7-97b3-657a5ba6c742");
        private static readonly Guid RatingsFieldGuid_RatingCount = new Guid("b1996002-9167-45e5-a4df-b2c41c6723c7");
        private static readonly Guid RatingsFieldGuid_RatedBy = new Guid("4D64B067-08C3-43DC-A87B-8B8E01673313");
        private static readonly Guid RatingsFieldGuid_Ratings = new Guid("434F51FB-FFD2-4A0E-A03B-CA3131AC67BA");
        private static readonly Guid LikeFieldGuid_LikedBy = new Guid("2cdcd5eb-846d-4f4d-9aaf-73e8e73c7312");
        private static readonly Guid LikeFieldGuid_LikeCount = new Guid("6e4d832b-f610-41a8-b3e0-239608efda41");

        private static List _library;

        #endregion

        /// <summary>
        /// Enable Social Settings Likes/Ratings on list. 
        /// Note: 1. Requires Publishing feature enabled on the web.
        ///       2. Defaults enable Ratings Experience on the List
        ///       3. When experience set to None, fields are not removed from the list since CSOM does not support removing hidden fields
        /// </summary>
        /// <param name="list">Current List</param>
        /// <param name="experience">Likes/Ratings</param>
        public static void SetRating(this List list, VotingExperience experience)
        {
            /*  Validate if current web is publishing web
             *  Add property to RootFolder of List/Library : key: Ratings_VotingExperience value:Likes
             *  Add rating fields
             *  Add fields to default view
             * */

            _library = list;

            //  only process publishing web
            if (!list.ParentWeb.IsPublishingWeb())
            {
                ////_logger.WriteWarning("Is publishing site : No");
            }

            //  Add to property Root Folder of Pages Library
            SetProperty(experience);

            if (experience == VotingExperience.None)
            {
                RemoveViewFields();
            }
            else
            {
                AddListFields();
                // remove already existing fields from the default view to prevent duplicates
                RemoveViewFields(); 
                AddViewFields(experience);
            }

            //_logger.WriteSuccess(string.Format("Enabled {0} on list/library {1}", experience, _library.Title));
        }

        /// <summary>
        /// Removes the view fields associated with any rating type
        /// </summary>
        private static void RemoveViewFields()
        {
            _library.Context.Load(_library.DefaultView, p => p.ViewFields);
            _library.Context.ExecuteQueryRetry();

            var defaultView = _library.DefaultView;

            if (defaultView.ViewFields.Contains("LikesCount"))
                defaultView.ViewFields.Remove("LikesCount");
            if (defaultView.ViewFields.Contains("AverageRating"))
                defaultView.ViewFields.Remove("AverageRating");

            defaultView.Update();
            _library.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Add Ratings/Likes related fields to List from current Web
        /// </summary>
        private static void AddListFields()
        {
            EnsureField(_library, RatingsFieldGuid_RatingCount);
            EnsureField(_library, RatingsFieldGuid_RatedBy);
            EnsureField(_library, RatingsFieldGuid_Ratings);
            EnsureField(_library, RatingsFieldGuid_AverageRating);
            EnsureField(_library, LikeFieldGuid_LikedBy);
            EnsureField(_library, LikeFieldGuid_LikeCount);

            _library.Update();
            _library.Context.ExecuteQueryRetry();
        }

        /// <summary>
        /// Add/Remove Ratings/Likes field in default view depending on exerpeince selected
        /// </summary>
        /// <param name="experience"></param>
        private static void AddViewFields(VotingExperience experience)
        {
            //  Add LikesCount and LikeBy (Explicit) to view fields
            _library.Context.Load(_library.DefaultView, p => p.ViewFields);
            _library.Context.ExecuteQueryRetry();

            var defaultView = _library.DefaultView;

            switch (experience)
            {
                case VotingExperience.Ratings:
                    //  Remove Likes Fields
                    if (defaultView.ViewFields.Contains("LikesCount"))
                        defaultView.ViewFields.Remove("LikesCount");

                    defaultView.ViewFields.Add("AverageRating");
                    //  Add Ratings related field
                    break;
                case VotingExperience.Likes:
                    //  Remove Ratings Fields
                    //  Add Likes related field
                    if (defaultView.ViewFields.Contains("AverageRating"))
                        defaultView.ViewFields.Remove("AverageRating");

                    defaultView.ViewFields.Add("LikesCount");
                    break;
                default:
                    throw new ArgumentOutOfRangeException("experience");
            }

            defaultView.Update();
            _library.Context.ExecuteQueryRetry();
            //_logger.WriteSuccess("Ensured view-field.");

        }

        /// <summary>
        /// Removes a field from a list
        /// </summary>
        /// <param name="list"></param>
        /// <param name="fieldId"></param>
        private static void RemoveField(List list, Guid fieldId)
        {
            try
            {
                FieldCollection listFields = list.Fields;
                Field field = listFields.GetById(fieldId);
                field.DeleteObject();
                _library.Context.ExecuteQueryRetry();
            }
            catch (ArgumentException)
            {
                // Field does not exist
            }
        }

        /// <summary>
        /// Check for Ratings/Likes field and add to ListField if doesn't exists.
        /// </summary>
        /// <param name="list">List</param>
        /// <param name="fieldId">Field Id</param>
        /// <returns></returns>
        private static void EnsureField(List list, Guid fieldId)
        {
            FieldCollection fields = list.Fields;

            FieldCollection availableFields = list.ParentWeb.AvailableFields;
            Field field = availableFields.GetById(fieldId);

            _library.Context.Load(fields);
            _library.Context.Load(field, p => p.SchemaXmlWithResourceTokens, p => p.Id, p => p.InternalName, p => p.StaticName);
            _library.Context.ExecuteQueryRetry();

            if (!fields.Any(p => p.Id == fieldId))
            {
                fields.AddFieldAsXml(field.SchemaXmlWithResourceTokens, false, AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToAllContentTypes);
            }
        }

        /// <summary>
        /// Add required key/value settings on List Root-Folder
        /// </summary>
        /// <param name="experience"></param>
        private static void SetProperty(VotingExperience experience)
        {
            _library.Context.Load(_library.RootFolder, p => p.Properties);
            _library.Context.ExecuteQueryRetry();

            if (experience != VotingExperience.None)
            { 
                _library.RootFolder.Properties["Ratings_VotingExperience"] = experience.ToString();
            }
            else
            {
                _library.RootFolder.Properties["Ratings_VotingExperience"] = string.Empty;
            }
            _library.RootFolder.Update();
            _library.Context.ExecuteQueryRetry();
            //_logger.WriteSuccess(string.Format("Ensured {0} Property.", experience));
        }
    }
}
