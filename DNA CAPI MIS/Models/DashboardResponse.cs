 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DNA_CAPI_MIS.Models
{
    public class DashboardResponse
    {

        public string Message { get; set; }
        public SurveyorStats All { get; set; }
        public SurveyorStats RHS { get; set; }
        public SurveyorStats MSU { get; set; }
        public SurveyorStats FWC { get; set; }
        public List<BarChart> OpenClose { get; set; }
        public List<Grid3> Grid3 { get; set; }
    }
}