using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Configuration;

using System.Linq.Expressions;
using System.Net;

namespace Adil.DAL
{
    /// <summary>
    /// The SharePointClient class is intended to encapsulate high performance, scalable best practices for
    /// SharePoint remote access
    /// </summary>
    public sealed class SharePointClient : IDisposable
    {
        #region Fields

        private SharePointLoginInfo _spLoginInfo;
        private ClientContext _ctx;

        #endregion

        #region Constructor

        /// <summary>
        /// Construct the SharePointClient object using login details
        /// </summary>
        /// <param name="spLoginInfo">An object of SharePointLoginInfo</param>
        public SharePointClient(SharePointLoginInfo spLoginInfo)
        {
            if (spLoginInfo.SiteURL == null)
                throw new ArgumentException("SharePoint site URL cannot be null/empty, its required to open connection.");

            this._spLoginInfo = spLoginInfo;

            this._ctx = new ClientContext(this._spLoginInfo.SiteURL);
            if (!string.IsNullOrEmpty(this._spLoginInfo.UserName))
            {
                this._ctx.Credentials = new NetworkCredential(this._spLoginInfo.DecryptedUserName, this._spLoginInfo.DecryptedPassword, this._spLoginInfo.Domain);
            }
        }

        /// <summary>
        /// Construct the SharePointClient object SharePoint ClientContext
        /// </summary>
        /// <param name="spClientContext">An object of ClientContext</param>
        public SharePointClient(ClientContext spClientContext)
        {

            this._ctx = spClientContext;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Get refrence of SharePoint ClientContext obejct
        /// </summary>
        public ClientContext SPClientContext
        {
            get
            {
                return this._ctx;
            }
        }

        #endregion

        #region Public Method

        public User GetUserInfo(out string picURL, out string myLocation, out string myBrand)
        {
            picURL = string.Empty;
            myLocation = ConfigurationManager.AppSettings[Constant.Configuration.NeonDefaultRegion];
            myBrand = ConfigurationManager.AppSettings[Constant.Configuration.NeonDefaultBrand];

            User currUser = this._ctx.Web.CurrentUser;
            this._ctx.Load(currUser, u => u.UserId, u => u.Email, u => u.Title);

            PeopleManager peopleManager = new PeopleManager(this._ctx);
            ClientResult<string> spLocation = peopleManager.GetUserProfilePropertyFor(this._spLoginInfo.UserLogin, Constant.SPUserLocationPropertyName);
            ClientResult<string> spPictureUrl = peopleManager.GetUserProfilePropertyFor(this._spLoginInfo.UserLogin, Constant.SPUserPictureURLPropertyName);
            ClientResult<string> spUserADPath = peopleManager.GetUserProfilePropertyFor(this._spLoginInfo.UserLogin, Constant.SPUserDistinguishedNamePropertyName);
            this._ctx.ExecuteQuery();

            picURL = spPictureUrl.Value;
            myLocation = string.IsNullOrEmpty(spLocation.Value) ? myLocation : spLocation.Value;
            //myADPath = spUserADPath.Value;
            return currUser;
        }

        public Uri GetUserPictureURL(string accountName, out string displayName)
        {
            PeopleManager peopleManager = new PeopleManager(this._ctx);
            PersonProperties props = peopleManager.GetPropertiesFor(accountName);
            this._ctx.Load(props, p => p.PictureUrl, p => p.DisplayName);
            this._ctx.ExecuteQuery();

            displayName = props.DisplayName;

            if (string.IsNullOrEmpty(props.PictureUrl))
                return null;
            else
                return new Uri(props.PictureUrl);
        }

        public ListItem GetItemByID(string webURL, string listURL, int listItemID, string[] fields)
        {
            List list = this._ctx.Site.OpenWeb(webURL).Lists.GetByTitle(listURL);
            ListItem itemById = list.GetItemById(listItemID);

            if (fields != null && fields.Length > 0)
            {
                Expression<Func<ListItem, object>>[] getSelectedFields = CreateListItemLoadExpressions(fields);

                this._ctx.Load(itemById, getSelectedFields);
            }
            else
            {
                this._ctx.Load(itemById, item => item.Id);
            }

            this._ctx.ExecuteQuery();

            return itemById;
        }

        public ListItem GetItemByID(Guid listID, int listItemID, string[] fields)
        {
            List list;

            if (this._spLoginInfo.WebID == Guid.Empty)
                list = this._ctx.Web.Lists.GetById(listID);
            else
                list = this._ctx.Site.OpenWebById(this._spLoginInfo.WebID).Lists.GetById(listID);

            ListItem itemById = list.GetItemById(listItemID);

            if (fields != null && fields.Length > 0)
            {
                Expression<Func<ListItem, object>>[] getSelectedFields = CreateListItemLoadExpressions(fields);

                this._ctx.Load(itemById, getSelectedFields);
            }
            else
            {
                //this._ctx.Load(itemById, item => item.Id);
                this._ctx.Load(itemById);
            }

            this._ctx.ExecuteQuery();

            return itemById;
        }

        public ListItemCollection GetItemByQuery(string webURL, string listURL, string viewFieldQuery, string[] fields)
        {
            List list = this._ctx.Site.OpenWeb(webURL).Lists.GetByTitle(listURL);
            CamlQuery query = new CamlQuery();
            query.ViewXml = string.IsNullOrEmpty(viewFieldQuery) ? "<View/>" : viewFieldQuery;

            ListItemCollection listItems = list.GetItems(query);

            Expression<Func<ListItemCollection, object>>[] getSelectedFields = CreateListItemCollectionLoadExpressions(fields);

            this._ctx.Load(listItems, getSelectedFields);

            this._ctx.ExecuteQuery();

            return listItems;
        }

        public ListItemCollection GetItemByQuery(Guid listID, string viewFieldQuery, string[] fields)
        {
            List list;

            if (this._spLoginInfo.WebID == Guid.Empty)
                list = this._ctx.Web.Lists.GetById(this._spLoginInfo.ListID);
            else
                list = this._ctx.Site.OpenWebById(this._spLoginInfo.WebID).Lists.GetById(listID);

            CamlQuery query = new CamlQuery();
            query.ViewXml = string.IsNullOrEmpty(viewFieldQuery) ? "<View/>" : viewFieldQuery;

            ListItemCollection listItems = list.GetItems(query);

            Expression<Func<ListItemCollection, object>>[] getSelectedFields = CreateListItemCollectionLoadExpressions(fields);

            this._ctx.Load(listItems, getSelectedFields);

            this._ctx.ExecuteQuery();

            return listItems;
        }

        public ResultTable Search(string keywords, List<string> properties, Dictionary<string, bool> sorting, Guid sourceID, int startRowIndex, int rowLimit)
        {
            KeywordQuery keywordQuery = new KeywordQuery(this._ctx);
            keywordQuery.TrimDuplicates = false;
            keywordQuery.QueryText = keywords;
            keywordQuery.SourceId = sourceID;
            keywordQuery.StartRow = startRowIndex;
            keywordQuery.RowsPerPage = rowLimit;
            keywordQuery.RowLimit = rowLimit;
            keywordQuery.SelectProperties.Clear();

            if (properties != null && properties.Count > 0)
            {
                foreach (string p in properties)
                    keywordQuery.SelectProperties.Add(p);
            }

            if (sorting != null && sorting.Count > 0)
            {
                foreach (var sort in sorting)
                {
                    if (sort.Value)
                        keywordQuery.SortList.Add(sort.Key, SortDirection.Ascending);
                    else
                        keywordQuery.SortList.Add(sort.Key, SortDirection.Descending);
                }
            }

            SearchExecutor searchExecutor = new SearchExecutor(this._ctx);
            ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
            this._ctx.ExecuteQuery();

            return results.Value[0];

        }

        #endregion

        #region Private Methods

        private Expression<Func<ListItemCollection, object>>[] CreateListItemCollectionLoadExpressions(string[] viewFields)
        {
            List<Expression<Func<ListItemCollection, object>>> expressions = new List<Expression<Func<ListItemCollection, object>>>();

            foreach (string viewFieldEntry in viewFields)
            {
                string strViewFieldEntry = viewFieldEntry;
                Expression<Func<ListItemCollection, object>> retrieveFieldDataExpression = listItems => listItems.Include(item => item[strViewFieldEntry]);
                expressions.Add(retrieveFieldDataExpression);
            }

            return expressions.ToArray();
        }

        private Expression<Func<ListItem, object>>[] CreateListItemLoadExpressions(string[] viewFields)
        {
            List<Expression<Func<ListItem, object>>> expressions = new List<Expression<Func<ListItem, object>>>();

            foreach (string viewFieldEntry in viewFields)
            {
                string strViewFieldEntry = viewFieldEntry;
                Expression<Func<ListItem, object>> retrieveFieldDataExpression = listItem => listItem[strViewFieldEntry];
                expressions.Add(retrieveFieldDataExpression);
            }

            return expressions.ToArray();
        }

        #endregion

        #region Disposing

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (this._ctx != null)
                {
                    this._ctx.Dispose();
                    this._ctx = null;
                }
            }
        }

        #endregion
    }

   

}
