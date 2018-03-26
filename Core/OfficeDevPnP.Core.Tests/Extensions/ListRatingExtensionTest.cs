using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Tests;

namespace Microsoft.SharePoint.Client.Tests
{
    [TestClass()]
    public class ListRatingExtensionTest
    {
        // Publishing infra active
        private static bool publishingActive = false;

        private const string Averagerating = "AverageRating";
        private const string Ratedby = "RatedBy";
        private const string Ratingcount = "RatingCount";
        private const string Likescount = "LikesCount";
        private const string Ratings = "Ratings";
        private const string Likedby = "LikedBy";
        private const string RatingsVotingexperience = "Ratings_VotingExperience";

        private ClientContext _clientContext;
        private List _list;

        #region Test initialize and cleanup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            using (ClientContext cc = TestCommon.CreateClientContext())
            {
                publishingActive = cc.Web.IsPublishingWeb();

                // Activate publishing
                if (!publishingActive)
                {
                    if (!cc.Site.IsFeatureActive(Constants.FeatureId_Site_Publishing))
                    {
                        cc.Site.ActivateFeature(Constants.FeatureId_Site_Publishing);
                    }

                    if (!cc.Web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                    {
                        cc.Web.ActivateFeature(Constants.FeatureId_Web_Publishing);
                    }
                }
            }

        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            using (ClientContext cc = TestCommon.CreateClientContext())
            {
                // deactivate publishing if it was not active before the test run
                if (!publishingActive)
                {
                    if (cc.Web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                    {
                        cc.Web.DeactivateFeature(Constants.FeatureId_Web_Publishing);
                    }

                    if (cc.Site.IsFeatureActive(Constants.FeatureId_Site_Publishing))
                    {
                        cc.Site.DeactivateFeature(Constants.FeatureId_Site_Publishing);
                    }
                }
            }
        }

        [TestInitialize()]
        public void Initialize()
        {
            /*** Make sure that the user defined in the App.config has permissions to Manage Terms ***/
            _clientContext = TestCommon.CreateClientContext();

            // Create Simple List
            _list = _clientContext.Web.CreateList(ListTemplateType.Contacts, "Test_list_" + DateTime.Now.ToFileTime(), false);
            _clientContext.Load(_list);
            _clientContext.ExecuteQueryRetry();

        }

        [TestCleanup]
        public void Cleanup()
        {
            // Clean up list
            _list.DeleteObject();
            _clientContext.ExecuteQueryRetry();
        }
        #endregion

        #region Rating's Test Scenarios
        [TestMethod()]
        public void EnableRatingExperienceTest()
        {
            // Enable Publishing Feature on Site and Web 

            if (!_clientContext.Site.IsFeatureActive(Constants.FeatureId_Site_Publishing))
                _clientContext.Site.ActivateFeature(Constants.FeatureId_Site_Publishing);

            if (!_clientContext.Web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                _clientContext.Web.ActivateFeature(Constants.FeatureId_Web_Publishing);

            _list.SetRating(VotingExperience.Ratings);

            //  Check if the Rating Fields are added to List, Views and Root Folder Property 

            Assert.IsTrue(IsRootFolderPropertySet(), "Root Folder property not set");
            Assert.IsTrue(HasRatingFields(), "Missing Rating Fields in List.");
            Assert.IsTrue(RatingFieldSetOnDefaultView(), "Required rating fields not added to default view.");

        }

        [TestMethod()]
        public void EnableLikesExperienceTest()
        {
            // Enable Publishing Feature on Site and Web 

            if (!_clientContext.Site.IsFeatureActive(Constants.FeatureId_Site_Publishing))
                _clientContext.Site.ActivateFeature(Constants.FeatureId_Site_Publishing);

            if (!_clientContext.Web.IsFeatureActive(Constants.FeatureId_Web_Publishing))
                _clientContext.Web.ActivateFeature(Constants.FeatureId_Web_Publishing);

            _list.SetRating(VotingExperience.Likes);

            //  Check if the Rating Fields are added to List, Views and Root Folder Property 

            Assert.IsTrue(IsRootFolderPropertySet(VotingExperience.Likes), "Required Root Folder property not set.");
            Assert.IsTrue(HasRatingFields(), "Missing Rating Fields in List.");
            Assert.IsTrue(RatingFieldSetOnDefaultView(VotingExperience.Likes), "Required rating fields not added to default view.");

        }

        #endregion


        /// <summary>
        /// Validate if required experience selected fields are added to default view
        /// </summary>
        /// <param name="experience"></param>
        /// <returns></returns>
        private bool RatingFieldSetOnDefaultView(VotingExperience experience = VotingExperience.Ratings)
        {
            _clientContext.Load(_list.DefaultView.ViewFields);
            _clientContext.ExecuteQueryRetry();

            switch (experience)
            {
                case VotingExperience.Ratings:
                    return _list.DefaultView.ViewFields.Contains(Averagerating);
                case VotingExperience.Likes:
                    return _list.DefaultView.ViewFields.Contains(Likescount);
                default:
                    throw new ArgumentOutOfRangeException("experience");
            }
        }

        /// <summary>
        /// Validates if required rating fields are present in the list.
        /// </summary>
        /// <returns></returns>
        private bool HasRatingFields()
        {
            _clientContext.Load(_list.Fields, p => p.Include(prop => prop.InternalName));
            _clientContext.ExecuteQueryRetry();

            var avgRating = _list.Fields.FirstOrDefault(p => p.InternalName == Averagerating);
            var ratedBy = _list.Fields.FirstOrDefault(p => p.InternalName == Ratedby);
            var ratingCount = _list.Fields.FirstOrDefault(p => p.InternalName == Ratingcount);
            var likeCount = _list.Fields.FirstOrDefault(p => p.InternalName == Likescount);
            var ratings = _list.Fields.FirstOrDefault(p => p.InternalName == Ratings);
            var likedBy = _list.Fields.FirstOrDefault(p => p.InternalName == Likedby);

            var fieldsExist = avgRating != null && ratedBy != null && ratingCount != null && ratings != null &&
                              likeCount != null && likedBy != null;

            return fieldsExist;
        }


        /// <summary>
        /// Validate if the RootFolder property is set appropriately
        /// </summary>
        /// <returns></returns>
        private bool IsRootFolderPropertySet(VotingExperience experience = VotingExperience.Ratings)
        {
            _clientContext.Load(_list.RootFolder.Properties);
            _clientContext.ExecuteQueryRetry();

            if (_list.RootFolder.Properties.FieldValues.ContainsKey(RatingsVotingexperience))
            {
                object exp;
                if (_list.RootFolder.Properties.FieldValues.TryGetValue(RatingsVotingexperience, out exp))
                {
                    return string.Compare(exp.ToString(), experience.ToString(), StringComparison.InvariantCultureIgnoreCase) == 0;
                }
            }

            return false;
        }



    }
}