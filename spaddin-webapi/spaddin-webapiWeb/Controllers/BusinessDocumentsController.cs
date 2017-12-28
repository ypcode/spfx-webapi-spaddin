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
        public static List<BusinessDocumentViewModel> MockBusinessDocuments = new List<BusinessDocumentViewModel>()
        {
            new BusinessDocumentViewModel()
            {
                Id = 1,
                Name = "Sample Document.docx",
                Purpose = "Agreement project"
            },
            new BusinessDocumentViewModel()
            {
                Id = 2,
                Name = "Sample Document 2.docx",
                Purpose = "Agreement project"
            },
            new BusinessDocumentViewModel()
            {
                Id = 3,
                Name = "Sample Document 3.docx",
                Purpose = "Offer project"
            },
            new BusinessDocumentViewModel()
            {
                Id = 4,
                Name = "Sample Document 4.docx",
                Purpose = "Research document"
            }
        };
        
        // GET: api/BusinessDocuments
        public IEnumerable<BusinessDocumentViewModel> Get()
        {
            return MockBusinessDocuments.ToList();
        }

        // GET: api/BusinessDocuments/5
        public BusinessDocumentViewModel Get(int id)
        {
            return MockBusinessDocuments.FirstOrDefault(d => d.Id == id);
        }

        // POST: api/BusinessDocuments
        public IHttpActionResult Post([FromBody]BusinessDocumentViewModel value)
        {
            int newId = MockBusinessDocuments.Count + 1;
            value.Id = newId;
            MockBusinessDocuments.Add(value);
            return Created($"/api/BusinessDocuments/{newId}", value);
        }

        // PUT: api/BusinessDocuments/5
        public IHttpActionResult Put(int id, [FromBody]BusinessDocumentViewModel value)
        {
            var found = MockBusinessDocuments.FirstOrDefault(b => b.Id == id);
            if (found == null)
                return NotFound();

            found.Name = value.Name;
            found.Purpose = value.Purpose;

            return Ok();
        }

        // DELETE: api/BusinessDocuments/5
        public IHttpActionResult Delete(int id)
        {
            var found = MockBusinessDocuments.FirstOrDefault(b => b.Id == id);
            if (found == null)
                return NotFound();

            MockBusinessDocuments.Remove(found);
            return Ok();
        }
    }
}
