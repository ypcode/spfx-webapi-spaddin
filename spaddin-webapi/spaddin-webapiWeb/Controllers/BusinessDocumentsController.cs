using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.WebAPI;
using spaddin_webapiWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace spaddin_webapiWeb.Controllers
{
    [WebAPIContextFilter]
    public class BusinessDocumentsController : ApiController
    {
        public const string FileLeafRefField = "FileLeafRef";
        public const string InChargeField = "InCharge";
        public const string DocumentPurposeField = "DocumentPurpose";

        public static string[] ValidDocumentPurposes = new string[]
        {
            "Agreement project",
            "Offer project",
            "Purchase project",
            "Research document"
        };

        private static readonly Dictionary<int, string> _usersLoginCache = new Dictionary<int, string>();

        private static string GetInChargeUserLoginName(ListItem businessDocListItem)
        {
            FieldUserValue inChargeUserValue = businessDocListItem[InChargeField] as FieldUserValue;
            string inChargeValue = inChargeUserValue != null ? inChargeUserValue.LookupValue : string.Empty;

            if (!_usersLoginCache.ContainsKey(inChargeUserValue.LookupId))
            {
                ClientContext clientContext = businessDocListItem.Context as ClientContext;
                if (clientContext != null)
                {
                    User user = clientContext.Web.EnsureUser(inChargeUserValue.LookupValue);
                    clientContext.Load(user);
                    clientContext.ExecuteQuery();
                    _usersLoginCache.Add(inChargeUserValue.LookupId, user.LoginName);
                }
            }

            return _usersLoginCache[inChargeUserValue.LookupId];
        }
        private static BusinessDocumentViewModel ListItemToViewModel(ListItem businessDocListItem)
        {
            return new BusinessDocumentViewModel()
            {
                Id = businessDocListItem.Id,
                Name = (string)businessDocListItem[FileLeafRefField],
                Purpose = (string)businessDocListItem[DocumentPurposeField],
                InCharge = GetInChargeUserLoginName(businessDocListItem)
            };
        }

        private static ListItem MapToListItem(BusinessDocumentViewModel viewModel, ListItem targetListItem)
        {
            targetListItem[FileLeafRefField] = viewModel.Name;
            targetListItem[DocumentPurposeField] = viewModel.Purpose;
            targetListItem[InChargeField] = FieldUserValue.FromUser(viewModel.InCharge);
            return targetListItem;
        }

        private static ListItem TryGetListItemById(List list, int id)
        {
            try
            {
                ListItem item = list.GetItemById(id);
                list.Context.Load(item, i => i.Id);
                list.Context.ExecuteQuery();
                return item;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool ValidateModel(BusinessDocumentViewModel viewModel, out string message)
        {
            // Validate the purpose is a valid value
            if (!ValidDocumentPurposes.Contains(viewModel.Purpose))
            {
                message = "The specified document purpose is invalid";
                return false;
            }

            message = string.Empty;
            return true;
        }

        // GET: api/BusinessDocuments
        public IEnumerable<BusinessDocumentViewModel> Get()
        {
            using (var clientContext = WebAPIHelper.GetClientContext(this.ControllerContext))
            {
                // Get the documents from the Business Documents library
                List businessDocsLib = clientContext.Web.GetListByUrl("/BusinessDocs");
                ListItemCollection businessDocItems = businessDocsLib.GetItems(CamlQuery.CreateAllItemsQuery());

                clientContext.Load(businessDocItems, items => items.Include(
                    item => item.Id,
                    item => item[FileLeafRefField],
                    item => item[InChargeField],
                    item => item[DocumentPurposeField]));
                clientContext.ExecuteQuery();

                // Create collection of view models from list item collection
                List<BusinessDocumentViewModel> viewModels = businessDocItems.Select(ListItemToViewModel).ToList();

                return viewModels;
            }
        }

        // GET: api/MyBusinessDocuments
        [HttpGet]
        [Route("api/MyBusinessDocuments")]
        public IEnumerable<BusinessDocumentViewModel> MyBusinessDocuments()
        {
            using (var clientContext = WebAPIHelper.GetClientContext(this.ControllerContext))
            {

                // Get the documents from the Business Documents library
                List businessDocsLib = clientContext.Web.GetListByUrl("/BusinessDocs");
                var camlQuery = new CamlQuery
                {
                    ViewXml = $@"<View><Query><Where>
    <Eq>
        <FieldRef Name='{InChargeField}' LookupId='TRUE' />
        <Value Type = 'Integer'><UserID /></Value>
     </Eq>
 </Where></Query></View>"
                };
                ListItemCollection businessDocItems = businessDocsLib.GetItems(camlQuery);
                
                clientContext.Load(businessDocItems, items => items.Include(
                    item => item.Id,
                    item => item[FileLeafRefField],
                    item => item[InChargeField],
                    item => item[DocumentPurposeField]));
                clientContext.ExecuteQuery();

                // Create collection of view models from list item collection
                List<BusinessDocumentViewModel> viewModels = businessDocItems.Select(ListItemToViewModel).ToList();

                return viewModels;
            }
        }

        

        // GET: api/BusinessDocuments/5
        public IHttpActionResult Get(int id)
        {
            using (var clientContext = WebAPIHelper.GetClientContext(this.ControllerContext))
            {
                // Get the documents from the Business Documents library
                List businessDocsLib = clientContext.Web.GetListByUrl("/BusinessDocs");
                ListItem businessDocItem = TryGetListItemById(businessDocsLib, id);
                if (businessDocItem == null)
                    return NotFound();

                // Ensure the needed metadata are loaded
                clientContext.Load(businessDocItem, item => item.Id,
                    item => item[FileLeafRefField],
                    item => item[InChargeField],
                    item => item[DocumentPurposeField]);
                clientContext.ExecuteQuery();

                // Create a view model object from the list item
                BusinessDocumentViewModel viewModel = ListItemToViewModel(businessDocItem);

                return Ok(viewModel);
            }
        }

        // POST: api/BusinessDocuments
        public IHttpActionResult Post([FromBody]BusinessDocumentViewModel value)
        {
            string validationError = null;
            if (!ValidateModel(value, out validationError))
            {
                return BadRequest(validationError);
            }

            using (var clientContext = WebAPIHelper.GetClientContext(this.ControllerContext))
            {
                // Get the documents from the Business Documents library
                List businessDocsLib = clientContext.Web.GetListByUrl("/BusinessDocs");
                // Ensure the root folder is loaded
                Folder rootFolder = businessDocsLib.EnsureProperty(l => l.RootFolder);
                ListItem newItem = businessDocsLib.CreateDocument(value.Name, rootFolder, DocumentTemplateType.Word);

                // Update the new document metadata
                newItem[DocumentPurposeField] = value.Purpose;
                newItem[InChargeField] = FieldUserValue.FromUser(value.InCharge);
                newItem.Update();

                // Ensure the needed metadata are loaded
                clientContext.Load(newItem, item => item.Id,
                    item => item[FileLeafRefField],
                    item => item[InChargeField],
                    item => item[DocumentPurposeField]);

                newItem.File.CheckIn("", CheckinType.MajorCheckIn);
                clientContext.ExecuteQuery();

                BusinessDocumentViewModel viewModel = ListItemToViewModel(newItem);
                return Created($"/api/BusinessDocuments/{viewModel.Id}", viewModel);
            }
        }

        // PUT: api/BusinessDocuments/5
        public IHttpActionResult Put(int id, [FromBody]BusinessDocumentViewModel value)
        {
            string validationError = null;
            if (!ValidateModel(value, out validationError))
            {
                return BadRequest(validationError);
            }

            using (var clientContext = WebAPIHelper.GetClientContext(this.ControllerContext))
            {
                // Get the documents from the Business Documents library
                List businessDocsLib = clientContext.Web.GetListByUrl("/BusinessDocs");
                ListItem businessDocItem = TryGetListItemById(businessDocsLib, id);

                // If not found, return the appropriate status code
                if (businessDocItem == null)
                    return NotFound();

                // Update the list item properties
                MapToListItem(value, businessDocItem);
                businessDocItem.Update();
                clientContext.ExecuteQuery();

                return Ok();
            }
        }

        // DELETE: api/BusinessDocuments/5
        public IHttpActionResult Delete(int id)
        {
            using (var clientContext = WebAPIHelper.GetClientContext(this.ControllerContext))
            {
                // Get the document from the Business Documents library
                List businessDocsLib = clientContext.Web.GetListByUrl("/BusinessDocs");
                ListItem businessDocItem = TryGetListItemById(businessDocsLib, id);
                // If not found, return the appropriate status code
                if (businessDocItem == null)
                    return NotFound();

                // Delete the list item
                businessDocItem.DeleteObject();
                clientContext.ExecuteQuery();

                return Ok();
            }
        }
    }
}
