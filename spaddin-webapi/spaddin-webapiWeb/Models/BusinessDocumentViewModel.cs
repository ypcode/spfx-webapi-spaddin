using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace spaddin_webapiWeb.Models
{
    public class BusinessDocumentViewModel
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Purpose { get; set; }

        public string InCharge { get; set; }
    }
}