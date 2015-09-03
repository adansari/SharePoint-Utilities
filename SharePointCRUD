using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;

namespace PublicisGroupe.EmailSignature.Common.Entities
{
    public class BackOffice : SPRepository, IDisposable
    {
        #region Fields

        private bool _isWebFromSPContext = false;
        private SPWeb _spweb;
        private List<string> _userGroups;

        #endregion Fields

        #region Constructors

        public BackOffice(SPWeb web)
        {
            this._spweb = web;
            this._userGroups = web.CurrentUser.Groups.Cast<SPGroup>().Select<SPGroup, string>(g => g.Name).ToList<string>();
        }

        public BackOffice(SPContext spContext)
        {
            this._spweb = spContext.Web;
            this._isWebFromSPContext = true;
            this._userGroups = spContext.Web.CurrentUser.Groups.Cast<SPGroup>().Select<SPGroup, string>(g => g.Name).ToList<string>();
        }

        #endregion Constructors

        #region Properties

        public ContentManagerType UserType
        {
            get
            {
                ContentManagerType type = ContentManagerType.Visitor;
                try
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate()
                    {
                        using (SPSite site = new SPSite(this._spweb.Url))
                        {
                            using (SPWeb elevatedWeb = site.OpenWeb())
                            {
                                List<SPGroup> cmGroups = new List<SPGroup>();

                                cmGroups = (from g in this._spweb.CurrentUser.Groups.Cast<SPGroup>()
                                            where g.Name.StartsWith(Constants.ContentMgr_Grp_Prefix) && Utility.CheckGroupHasRole(elevatedWeb, g, Constants.ContentManager_Perm_Name)
                                            select g).ToList<SPGroup>();

                                if (cmGroups != null && cmGroups.Count > 0)
                                {
                                    foreach (var grp in cmGroups)
                                    {
                                        string countryID = string.Empty;
                                        int[] dirIDs = Utility.GetDirectoryIDsFromContentMgrGroupName(grp.Name, out countryID);

                                        if (!string.IsNullOrEmpty(countryID))
                                        {
                                            if (dirIDs[0] > 0 && dirIDs[1] > 0)
                                            {
                                                type = type | ContentManagerType.BrandCountryAdmin;
                                            }
                                            else if (dirIDs[0] > 0)
                                            {
                                                type = type | ContentManagerType.ParentBrandCountryAdmin;
                                            }
                                            else
                                            {
                                                type = type | ContentManagerType.CountryAdmin;
                                            }
                                        }
                                        else
                                        {
                                            if (dirIDs[0] > 0 && dirIDs[1] > 0 && dirIDs[2] > 0)
                                            {
                                                type = type | ContentManagerType.AgencyAdmin;
                                            }
                                            else if (dirIDs[0] > 0 && dirIDs[1] > 0)
                                            {
                                                type = type | ContentManagerType.BrandAdmin;
                                            }
                                            else
                                            {
                                                type = type | ContentManagerType.ParentBrandAdmin;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    });
                }
                catch (Exception ex) { Logger.Write(this.WebURL,ex, ErroredModule.BackOffice); }

                if (this.HasUserFullControl)
                {
                    type = type | ContentManagerType.ApplicationAdmin;
                }

                return type;
            }
        }

        public string WebURL
        {
            get { return this._spweb.Url; }
        }

        #endregion Properties

        #region Methods

        public ParentBrand CreateParentBrand(string name, int gadID)
        {
            Dictionary<string, object> values = new Dictionary<string, object>();
            values.Add(Constants.Parent_Brand_Name_Field, name);
            values.Add(Constants.IsDeleted_Field, false);
            if (gadID > 0) values.Add(Constants.GAD_ID_Field, gadID.ToString());

            int id = this.AddNew(Constants.Parent_Brand_Url, values);

            return new ParentBrand(id, name, gadID, this);
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void GetAllBrandsAndAgencies(out IDirectoryCollection allBrands, out IDirectoryCollection allAgencies, bool getAgencies, bool includeDeleted, bool applySecurityTrimming = true)
        {
            allBrands = new IDirectoryCollection();
            allAgencies = new IDirectoryCollection();

            foreach (ParentBrand pb in this.GetChildren(includeDeleted, applySecurityTrimming))
            {
                IDirectoryCollection pbBrands = pb.GetChildren(includeDeleted, applySecurityTrimming);
                allBrands.AddRange(pbBrands);

                if (getAgencies)
                {
                    foreach (Brand b in pbBrands)
                    {
                        IDirectoryCollection bAgencies = b.GetChildren(includeDeleted, applySecurityTrimming);
                        allAgencies.AddRange(bAgencies);
                    }
                }

            }
        }

        public List<Template> GetAllTemplates()
        {
            return GetAllTemplates(string.Empty);
        }

        public IDirectoryCollection GetChildren(bool includeDeleted, bool applySecurityTrimming = true)
        {
            string viewFields = string.Concat(
                                   "<FieldRef Name='" + Constants.List_ID_Field + "' />",
                                   "<FieldRef Name='" + Constants.GAD_ID_Field + "' />",
                                   "<FieldRef Name='" + Constants.Notification_InheritParentNotification_Field + "' />",
                                   "<FieldRef Name='" + Constants.Field_InheritParentFields_Field + "' />",
                                   "<FieldRef Name='" + Constants.Notification_IsNotificationActive_Field + "' />",
                                   "<FieldRef Name='" + Constants.Parent_Brand_Name_Field + "' />",
                                   "<FieldRef Name='" + Constants.IsDeleted_Field + "' />"
                                   );

            string query=string.Empty;

            if (!includeDeleted)
                query = "<Where><Eq><FieldRef Name='" + Constants.IsDeleted_Field + "' /><Value Type='Boolean'>0</Value></Eq></Where>";

            SPListItemCollection spItems = GetItemByQuery(Constants.Parent_Brand_Url, viewFields, query);

            IDirectoryCollection parentBrands = new IDirectoryCollection();

            foreach (SPListItem item in spItems)
            {
                int id = Convert.ToInt32(item[Constants.List_ID_Field]);
                int gadID = Convert.ToInt32(item[Constants.GAD_ID_Field]);
                string name = Convert.ToString(item[Constants.Parent_Brand_Name_Field]);
                bool inheritParentField = Convert.ToBoolean(item[Constants.Field_InheritParentFields_Field] ?? "false");
                bool inheritParentNoti = Convert.ToBoolean(item[Constants.Notification_InheritParentNotification_Field] ?? "false");
                bool isaAtiveNotification = Convert.ToBoolean(item[Constants.Notification_IsNotificationActive_Field] ?? "false");
                bool isDeleted = Convert.ToBoolean(item[Constants.IsDeleted_Field] ?? "false");
                parentBrands.Add(new ParentBrand(id, name, gadID,this) { InheritParentFields = inheritParentField, InheritParentNotification = inheritParentNoti, IsNotificationActive = isaAtiveNotification,IsDeleted=isDeleted });
            }

            if (applySecurityTrimming)
            {
                IDirectoryCollection parentBrandsSecured = new IDirectoryCollection();
                parentBrandsSecured.AddRange(parentBrands.FindAll(pb => pb.DoesUserHasPermission));
                return parentBrandsSecured;
            }
            else
            {
                return parentBrands;
            }
        }

        public string GetChildrenAsXML()
        {
            IDirectoryCollection pbrands = this.GetChildren(false, false);

            string includeDirHavingAnyTemplate = "<FieldRef Name='ID' /><Values><Value Type='Number'>0</Value>";

            foreach (var pb in pbrands)
            {
                if (pb.GetAllTemplates().Count > 0 || pb.GetChildren(false, false).FindAll(br => br.GetAllTemplates().Count > 0 || br.GetChildren(false,false).FindAll(ag=>ag.GetAllTemplates().Count>0).Count>0).Count > 0)
                {
                    includeDirHavingAnyTemplate += string.Format("<Value Type='Number'>{0}</Value>", pb.ID);
                }
            }

            includeDirHavingAnyTemplate += "</Values>";

            string viewFields = string.Concat(
                                   "<FieldRef Name='" + Constants.List_ID_Field + "' />",
                                   "<FieldRef Name='" + Constants.GAD_ID_Field + "' />",
                                   "<FieldRef Name='" + Constants.Parent_Brand_Name_Field + "' />");

            string query = "<Where><And><Eq><FieldRef Name='" + Constants.IsDeleted_Field + "' /><Value Type='Boolean'>0</Value></Eq><In>" + includeDirHavingAnyTemplate + "</In></And></Where>";

            SPListItemCollection spItems = GetItemByQuery(Constants.Parent_Brand_Url, viewFields, query);
            return spItems.Xml;
        }

        public DataTable GetCountries()
        {
            string viewFields = string.Concat(
                                       "<FieldRef Name='" + Constants.List_ID_Field + "' />",
                                       "<FieldRef Name='" + Constants.Agency_CountryLabel_Field + "' />",
                                       "<FieldRef Name='" + Constants.Agency_Country_Field + "' />");

            string whereClause = "<OrderBy><FieldRef Name='" + Constants.Agency_CountryLabel_Field + "' Ascending='True' /></OrderBy>";

            SPListItemCollection result = this.GetItemByQuery(Constants.Agency_Url, viewFields, whereClause);

            if (result != null && result.Count > 0)
            {
                return result.GetDataTable().DefaultView.ToTable( /*distinct*/ true, new string[] { Constants.Agency_CountryLabel_Field,Constants.Agency_Country_Field });
            }
            else
            {
                return null;
            }
        }

        public DataTable SearchDirectoryWithCustomSettings(string nameFilter, ParentType dirType,bool includeDeleted)
        {
            IDirectoryCollection dirsWithCustomSettings = new IDirectoryCollection();
            IDirectoryCollection allBrands, allAgencies;
            string parentTypeName="Invalid";
            int parentBrandID = 0, brandID = 0, agencyID = 0; string parentBrandName = string.Empty, brandName = string.Empty, agencyName = string.Empty;
            this.GetAllBrandsAndAgencies(out allBrands, out allAgencies, true,includeDeleted);

            dirsWithCustomSettings.AddRange(this.GetChildren(includeDeleted).FindAll(dir => (!dir.InheritParentFields || dir.IsNotificationActive)));
            dirsWithCustomSettings.AddRange(allBrands.FindAll(dir => (!dir.InheritParentFields || !dir.InheritParentNotification)));
            dirsWithCustomSettings.AddRange(allAgencies.FindAll(dir => (!dir.InheritParentFields || !dir.InheritParentNotification)));

            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("ID", typeof(int));
            dtResult.Columns.Add("Name", typeof(string));
            dtResult.Columns.Add("ParentType", typeof(string));
            dtResult.Columns.Add("PBID", typeof(int));
            dtResult.Columns.Add("BID", typeof(int));
            dtResult.Columns.Add("AGID", typeof(int));
            dtResult.Columns.Add("IsNoticationActive", typeof(bool));
            dtResult.Columns.Add("InheritsNotification", typeof(bool));
            dtResult.Columns.Add("InheritParentFields", typeof(bool));
            dtResult.Columns.Add("PBName", typeof(string));
            dtResult.Columns.Add("BRName", typeof(string));
            dtResult.Columns.Add("AGName", typeof(string));

            var result = (from dir in dirsWithCustomSettings
                          where Utility.CheckDirType(dir, dirType, out parentTypeName, out parentBrandID, out brandID, out agencyID, out parentBrandName ,out brandName,out agencyName) &&
                               (string.IsNullOrEmpty(nameFilter) ? true : dir.Name.ToUpper().Contains(nameFilter.ToUpper())) &&
                               dir.DoesUserHasWritePermission
                          select new { 
                              ID = dir.ID, 
                              Name = dir.Name, 
                              ParentType = parentTypeName, 
                              PBID = parentBrandID, 
                              BID = brandID, 
                              AGID = agencyID, 
                              IsNoticationActive = dir.IsNotificationActive, 
                              InheritsNotification = dir.InheritParentNotification, 
                              InheritParentFields=dir.InheritParentFields,
                              PBName = parentBrandName,
                              BRName = brandName,
                              AGName = agencyName
                          }).ToList();


            foreach (var item in result)
            {
                dtResult.Rows.Add(item.ID,item.Name,item.ParentType,item.PBID,item.BID,item.AGID,item.IsNoticationActive,item.InheritsNotification,item.InheritParentFields,item.PBName,item.BRName,item.AGName);
            }

            return dtResult;
        }

        #endregion Methods

        #region SPRepository implementation

        #region Properties

        public override TemplateFields FieldSettings
        {
            get
            {
                return TemplateFields.GetApplicationFieldSettings(this);
            }
        }

        public override bool HasUserFullControl
        {
            get
            {
                return this._spweb.AllRolesForCurrentUser.Contains(this._spweb.RoleDefinitions[Constants.Admin_Perm_Name]);
            }
        }

        #endregion Properties

        #region Methods

        public override int AddNew(string listURL, Dictionary<string, object> values)
        {
            SPList list = this._spweb.Lists.TryGetList(listURL);

            if (list == null) throw new SPException(string.Format("List '{0}' does not exist.", listURL));

            SPListItem item = list.AddItem();

            if (values != null && values.Count > 0)
            {
                foreach (var entry in values)
                {
                    item[entry.Key] = entry.Value;
                }

                item.Update();
            }

            return item.ID;
        }

        public override bool CanUserReadDirectory(IDirectory dir)
        {
            bool canUserReadDir = false;

            foreach (string grpName in this._userGroups)
            {
                string countryID;
                int[] dirIDs = Utility.GetDirectoryIDsFromContentMgrGroupName(grpName, out countryID);

                if (string.IsNullOrEmpty(countryID))
                {
                    if (dir is ParentBrand)
                    {
                        if (dirIDs[0] == dir.ID) { canUserReadDir = true; break; }
                    }
                    else if (dir is Brand)
                    {
                        if (dirIDs[1] == dir.ID) { canUserReadDir = true; break; }
                    }
                    else if (dir is Agency)
                    {
                        if (dirIDs[2] == dir.ID) { canUserReadDir = true; break; }
                    }
                }
                else
                {
                    if (dir is ParentBrand)
                    {
                        if (dirIDs[0] == 0) { canUserReadDir = true; break; }
                        else
                        {
                            if (dirIDs[0] == dir.ID) { canUserReadDir = true; break; }
                        }
                    }
                    else if (dir is Brand)
                    {
                        if (dirIDs[1] == 0) { canUserReadDir = true; break; }
                        else
                        {
                            if (dirIDs[1] == dir.ID) { canUserReadDir = true; break; }
                        }
                    }
                    else if (dir is Agency)
                    {
                        if (((Agency)dir).Country.Trim().ToLower() == countryID.Trim().ToLower()) { canUserReadDir = true; break; }
                    }
                }
            }

            return canUserReadDir;
        }

        public override void DeleteByID(string listURL, int id)
        {
            SPList list = this._spweb.Lists.TryGetList(listURL);

            if (list == null) throw new SPException(string.Format("List '{0}' does not exist.", listURL));

            SPListItem itemToDelete = list.GetItemById(id);
            if (itemToDelete != null) itemToDelete.Recycle();
        }

        public override bool DoesUserBelongToGroup(string[] groups)
        {
            bool isUserBelongsToGroup = false;

            foreach (string g in groups)
            {
                if (this._userGroups.Contains(g))
                {
                    isUserBelongsToGroup = true;
                    break;
                }
            }

            return isUserBelongsToGroup;
        }

        public override bool DoesUserHasPermissionOnItem(SPListItem item, SPBasePermissions perms)
        {
            SPRoleDefinition roleDefinition = item.AllRolesForCurrentUser.Cast<SPRoleDefinition>().Where(x => x.Name.ToLower() == Constants.ContentManager_Perm_Name.ToLower() || x.Name.ToLower() == Constants.Admin_Perm_Name.ToLower()).FirstOrDefault();

            if (roleDefinition != null)
            {
                return true;
            }
            else
            {
                return false;
            }
            //return item.AllRolesForCurrentUser.Cast<SPRoleDefinition>().Any<SPRoleDefinition>(rd=> rd.Name.ToLower() == Constants.ContentManager_Perm_Name.ToLower() || rd.Name.ToLower() == Constants.Admin_Perm_Name.ToLower());

            //return item.DoesUserHavePermissions(this._spweb.CurrentUser, perms);
        }
        
        public override List<Template> GetAllTemplates(string whereClause)
        {
            string viewFields = string.Concat(
                                   "<FieldRef Name='" + Constants.List_ID_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Title_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Body_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_ParentBrand_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Brand_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Agency_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Agency_Field + "_x003a_" + Constants.Agency_Country_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_ParentType_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_IsOrphan_Field + "' />"
                                   );

            SPListItemCollection spItems = this.GetItemByQuery(Constants.Signature_Template_Url, viewFields, whereClause);

            List<Template> templates = new List<Template>();

            if (spItems == null) return templates;

            foreach (SPListItem item in spItems)
            {
                templates.Add(new Template(Convert.ToInt32(item[Constants.List_ID_Field]), this, false)
                {
                    Name = Convert.ToString(item[Constants.Signature_Template_Title_Field]),
                    Body = Convert.ToString(item[Constants.Signature_Template_Body_Field]),
                    ParentBrandValue = Convert.ToString(item[Constants.Signature_Template_ParentBrand_Field]),
                    BrandValue = Convert.ToString(item[Constants.Signature_Template_Brand_Field]),
                    AgencyValue = Convert.ToString(item[Constants.Signature_Template_Agency_Field]),
                    AgencyCountryValue = Convert.ToString(item[Constants.Signature_Template_Agency_Field + "_x003a_" + Constants.Agency_Country_Field]),
                    ParentType = Convert.ToString(item[Constants.Signature_Template_ParentType_Field]),
                    IsOrphan = Convert.ToBoolean(item[Constants.Signature_Template_IsOrphan_Field] ?? false)
        
                });
            }

            return templates;
        }
        
        public override List<Template> GetAllTemplatesWithPermissions(string whereClause)
        {
            string viewFields = string.Concat(
                                   "<FieldRef Name='" + Constants.List_ID_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Title_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Body_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_ParentBrand_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Brand_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Agency_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_Agency_Field + "_x003a_" + Constants.Agency_Country_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_ParentType_Field + "' />",
                                   "<FieldRef Name='" + Constants.Signature_Template_IsOrphan_Field + "' />"
                                   );

            SPListItemCollection oSpListItemColl = this.GetItemByQuery(Constants.Signature_Template_Url, viewFields, whereClause);

            List<Template> templates = new List<Template>();
            if (oSpListItemColl != null && oSpListItemColl.Count > 0)
            {
                foreach (SPListItem item in oSpListItemColl)
                {
                    bool userHasPErmission = this.DoesUserHasPermissionOnItem(item, SPBasePermissions.EditListItems);
                    templates.Add(new Template(Convert.ToInt32(item[Constants.List_ID_Field]), this,  false)
                    {
                        Name = Convert.ToString(item[Constants.Signature_Template_Title_Field]),
                        Body = Convert.ToString(item[Constants.Signature_Template_Body_Field]),
                        ParentBrandValue = Convert.ToString(item[Constants.Signature_Template_ParentBrand_Field]),
                        BrandValue = Convert.ToString(item[Constants.Signature_Template_Brand_Field]),
                        AgencyValue = Convert.ToString(item[Constants.Signature_Template_Agency_Field]),
                        AgencyCountryValue = Convert.ToString(item[Constants.Signature_Template_Agency_Field + "_x003a_" + Constants.Agency_Country_Field]),
                        ParentType = Convert.ToString(item[Constants.Signature_Template_ParentType_Field]),
                        IsOrphan = Convert.ToBoolean(item[Constants.Signature_Template_IsOrphan_Field] ?? false),
                        DoesUserHasWritePermission = userHasPErmission

                    });
                }
            }
            else
            {
                return templates;
            }

            return templates;
        }

        public override SPListItem GetItemByID(string listURL, int listItemID, string[] fields)
        {
            SPList list = this._spweb.Lists.TryGetList(listURL);

            if (list == null) throw new SPException(string.Format("List '{0}' does not exist.", listURL));

            SPListItem listItem = list.GetItemByIdSelectedFields(listItemID, fields);

            return listItem;
        }

        public override SPListItemCollection GetItemByQuery(string listURL, string viewFields, string whereClause)
        {
            SPList list = this._spweb.Lists.TryGetList(listURL);

            if (list == null) throw new SPException(string.Format("List '{0}' does not exist.", listURL));

            SPQuery query = new SPQuery();
            query.ViewFieldsOnly = true;

            if (string.IsNullOrEmpty(viewFields))
            {
                query.ViewFields = "<FieldRef Name='" + Constants.List_ID_Field + "' />";
            }
            else
            {
                query.ViewFields = viewFields;
            }

            query.Query = whereClause;

            return list.GetItems(query);
        }

        public override int Update(string listURL, int id, Dictionary<string, object> values)
        {
            SPList list = this._spweb.Lists.TryGetList(listURL);

            if (list == null) throw new SPException(string.Format("List '{0}' does not exist.", listURL));

            SPListItem item = list.GetItemByIdAllFields(id);

            foreach (var entry in values)
            {
                item[entry.Key] = entry.Value;
            }

            item.Update();

            return item.ID;
        }

        public override int UpdateByGADID(string listURL, int gadID, Dictionary<string, object> values)
        {
            SPList list = this._spweb.Lists.TryGetList(listURL);

            if (list == null) throw new SPException(string.Format("List '{0}' does not exist.", listURL));

            if (!list.Fields.ContainsField(Constants.GAD_ID_Field)) throw new SPException(string.Format("List '{0}' does not contains the field: {1}(int).", listURL, Constants.GAD_ID_Field));

            string gadWhereClause = "<Where><Eq><FieldRef Name='" + Constants.GAD_ID_Field + "' /><Value Type='Integer'>" + gadID + "</Value></Eq></Where>";

            SPListItemCollection items = this.GetItemByQuery(listURL, string.Empty, gadWhereClause);

            if (items.Count < 1) throw new SPException(string.Format("List '{0}' does not contains any item where {1}={2}", listURL, Constants.GAD_ID_Field, gadID));

            SPListItem item = list.GetItemByIdAllFields(items[0].ID);

            foreach (var entry in values)
            {
                item[entry.Key] = entry.Value;
            }

            item.Update();

            return item.ID;
        }

        public override List<string> GetDirectoryAdmin(int parentBrand, int brand, int agency,string country)
        {
            List<string> grpNames=new List<string>();
            List<string> adminEmails = new List<string>();

            grpNames.Add(Utility.GenerateContentMgrGroupName(parentBrand, brand, agency));

            if (!string.IsNullOrEmpty(country)) grpNames.AddRange(Utility.GeneratePossibleContentMgrGroupName(parentBrand, brand, country));

            if (grpNames.Count <= 0) return adminEmails;

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite(this._spweb.Url))
                {
                    using (SPWeb elevatedWeb = site.OpenWeb())
                    {
                        SPGroupCollection grpsFound = elevatedWeb.SiteGroups.GetCollection(grpNames.ToArray());

                        if (grpsFound.Count > 0)
                        {
                            foreach (SPGroup grp in grpsFound)
                            {
                                if (grp.Users.Count > 0)
                                    foreach (SPUser s in grp.Users) adminEmails.Add(s.Email);
                            }
                        }
                    }
                }
            });

            return adminEmails;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing && !this._isWebFromSPContext)
            {
                if (this._spweb != null) this._spweb.Dispose();
            }
        }

        #endregion Methods

        #endregion SPRepository implementation
    }
}
