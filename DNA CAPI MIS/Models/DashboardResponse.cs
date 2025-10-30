
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
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
        public SurveyorStats All { get; set; } = new SurveyorStats();
        public SurveyorStats RHS { get; set; } = new SurveyorStats();
        public SurveyorStats MSU { get; set; } = new SurveyorStats();
        public SurveyorStats FWC { get; set; } = new SurveyorStats();
        public List<PieChartOC> MSUOpenClose { get; set; } = new List<PieChartOC>();
        public List<PieChartOC> FWCOpenClose { get; set; } = new List<PieChartOC>();
        public List<PieChartOC> RHSOpenClose { get; set; } = new List<PieChartOC>();
        public List<Grid3> Grid3 { get; set; } = new List<Grid3>();
        public List<Grid14> FuniturePosition { get; set; } = new List<Grid14>();
        public int AllCnt { get; set; } = 0;
        public int RHSCnt { get; set; } = 0;
        public int MSUCnt { get; set; } = 0;
        public int FWCCnt { get; set; } = 0;


        public ContraceptiveStockPositionModel contraceptiveStockPositionModel { get; set; } = new ContraceptiveStockPositionModel();
    }


    public class SDPsStatus
    {
        public string asDate { get; set; }
        public string ProjectName { get; set; }
        public string District { get; set; }
        public string Center { get; set; }
        public string Premises { get; set; }
        public string OpenClose { get; set; }
        public string Status { get; set; }

    }
    public class StockOfContraceptiveResponse
    {
        public string asDate { get; set; }
        public string ProjectName { get; set; }
        public int sbjnum { get; set; }
        public int ProjectId { get; set; }
        public string StockOfContraceptive { get; set; }

    }

    public class Comodities
    {
        public string Condoms { get; set; }
        public string POP { get; set; }
        public string COC { get; set; }
        public string ECP { get; set; }
        public string ThreeMonth { get; set; }
        public string Defo { get; set; }
        public string IUD { get; set; }
        public string Jodelle { get; set; }
        public string Date { get; set; }
    }
}