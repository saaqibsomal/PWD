using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DNA_CAPI_MIS.Models
{
    public class DashboardRequest
    {
        public string StartDate { get;set; }
        public string EndDate { get;set; }
        public string DistrictID { get;set; }
        public string DistrictName { get;set; }
        public string CenterID { get;set; }
        public string CenterName { get;set; }
        public string ProjectId { get;set; }
        public string QuestionId { get;set; }
    }
}