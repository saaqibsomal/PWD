using DNA_CAPI_MIS.Models;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2010.Ink;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Org.BouncyCastle.Ocsp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Web.Mvc;

namespace DNA_CAPI_MIS.Service
{
    public class DashboardService
    {
        public DNA_CAPI_MIS.DAL.ProjectContext dbContext = new DNA_CAPI_MIS.DAL.ProjectContext();

        public dynamic Dashboard(DashboardRequest req)
        {
            DashboardResponse response = new DashboardResponse();
            MSUOpenClose(req, response);
            RHSOpenClose(req, response);
            FWCOpenClose(req, response);
            NumberOfVisitor(req, response);
            StatusOfBuilding(req, response);
            response.contraceptiveStockPositionModel = ContraceptiveStockPosition(req);
            response.FuniturePosition = FuniturePosition(req);
            return response;
        }

        public int GetFHSCenter()
        {
            var Central = dbContext.ProjectFieldSample.Where(x => x.IsActive && "50446".Contains(x.FieldID.ToString()));
            return 0;
        }
        public int GetMSUCenter()
        {
            var Central = dbContext.ProjectFieldSample.Where(x => x.IsActive && "50486".Contains(x.FieldID.ToString()));
            return 0;
        }
        public int GetFWCCenter()
        {
            var Central = dbContext.ProjectFieldSample.Where(x => x.IsActive && "55588".Contains(x.FieldID.ToString()));
            return 0;
        }
        private void NumberOfVisitor(DashboardRequest req, DashboardResponse response)
        {
            var All = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID in( 7120,7121,7122) and Convert(datetime, DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";
            var RHS = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID = 7120 and Convert(datetime, DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";
            var MSU = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID = 7121 and Convert(datetime, DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";
            var FWC = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID = 7122 and Convert(datetime, DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";

            var RHS_MSU_FWC = dbContext.Database.SqlQuery<SurveyorStats>(All);
            var queryRHS = dbContext.Database.SqlQuery<SurveyorStats>(RHS);
            var queryMSU = dbContext.Database.SqlQuery<SurveyorStats>(MSU);
            var queryFWC = dbContext.Database.SqlQuery<SurveyorStats>(FWC);

            if (RHS_MSU_FWC.Count() > 0)
            {
                response.All.SurveyorName = "All";
                response.All.SurveyCount = RHS_MSU_FWC.Sum(x => x.SurveyCount);
            }
            if (queryRHS.Count() > 0)
            {
                response.RHS.SurveyCount = queryRHS.Sum(x => x.SurveyCount);
            }
            response.RHS.SurveyorName = "RHS";
            if (queryMSU.Count() > 0)
            {
                response.MSU.SurveyCount = queryMSU.Sum(x => x.SurveyCount);
            }
            response.MSU.SurveyorName = "MSU";
            if (queryFWC.Count() > 0)
            {
                response.FWC.SurveyCount = queryFWC.Sum(x => x.SurveyCount);
            }
            response.FWC.SurveyorName = "FWC";
        }
        public void StatusOfBuilding(DashboardRequest req, DashboardResponse response)
        {
            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
    DROP TABLE #Graph;

;WITH cte AS (
    SELECT 
        s.ProjectID,
        s.sbjnum,
        s.DeviceTimestamp,
        sd2.FieldId AS FieldId2, sd2.FieldValue AS FieldValue2,
        sd3.FieldId AS FieldId3, sd3.FieldValue AS FieldValue3,
        sd5.FieldId AS FieldId5, sd5.FieldValue AS FieldValue5,
        sd6.FieldId AS FieldId6, sd6.FieldValue AS FieldValue6,
        sd7.FieldId AS FieldId7, sd7.FieldValue AS FieldValue7,
        DATENAME(MONTH, s.DeviceTimestamp) AS MonthName,
        MONTH(s.DeviceTimestamp) AS MonthNum,
        YEAR(s.DeviceTimestamp) AS YearNum,
        ROW_NUMBER() OVER (
            PARTITION BY sd3.FieldValue, YEAR(s.DeviceTimestamp), MONTH(s.DeviceTimestamp)   -- per center per month
            ORDER BY s.DeviceTimestamp ASC                                           -- first inserted record
        ) AS RowNum
    FROM Survey s
    INNER JOIN SurveyData sd2 ON s.sbjnum = sd2.sbjnum AND sd2.FieldId IN (50435, 50484, 55587) -- District
    INNER JOIN SurveyData sd3 ON s.sbjnum = sd3.sbjnum AND sd3.FieldId IN (50446, 50486, 55588) -- Center
    INNER JOIN SurveyData sd5 ON s.sbjnum = sd5.sbjnum AND sd5.FieldId IN (55594, 50498, 50461) -- Indicate Sign
    INNER JOIN SurveyData sd6 ON s.sbjnum = sd6.sbjnum AND sd6.FieldId IN (50557, 50500, 55595) -- Status of Building
    LEFT JOIN SurveyData sd7 ON s.sbjnum = sd7.sbjnum AND sd7.FieldId IN (50462, 50499, 55596)  -- Cleanliness
    WHERE s.ProjectID IN ({req.ProjectId})
)
SELECT  
    (SELECT TOP 1 p.[Name] FROM Project p WHERE p.Id = cte.ProjectID) AS ProjectName,
    fs2.Title AS District,
    fs3.Title AS Center,
    FieldValue5 AS IndicateSign,
    FieldValue6 AS StatusOfBuilding,
    CASE 
        WHEN FieldValue7 = '1' THEN 'Satisfactory'
        WHEN FieldValue7 = '2' THEN 'Not Satisfactory'
        ELSE ''
    END AS Cleanliness,
    CONVERT(VARCHAR, cte.DeviceTimestamp, 101) AS [asDate],
    cte.MonthName,
    cte.YearNum,
    cte.sbjnum,
    cte.MonthNum
INTO #Graph
FROM cte
INNER JOIN ProjectFieldSample fs2 ON cte.FieldId2 = fs2.FieldID AND fs2.Code IN (cte.FieldValue2)
INNER JOIN ProjectFieldSample fs3 ON cte.FieldId3 = fs3.FieldID AND fs3.Code IN (cte.FieldValue3)
LEFT JOIN ProjectFieldSample fs5 ON cte.FieldId5 = fs5.FieldID AND fs5.Code IN (cte.FieldValue5)
LEFT JOIN ProjectFieldSample fs6 ON cte.FieldId6 = fs6.FieldID AND fs6.Code IN (cte.FieldValue6)
LEFT JOIN ProjectFieldSample fs7 ON cte.FieldId7 = fs7.FieldID AND fs7.Code IN (cte.FieldValue7)
WHERE RowNum = 1
  AND CONVERT(DATETIME, cte.DeviceTimestamp, 101) 
      BETWEEN '{req.StartDate} 00:00:01' AND '{req.EndDate} 23:59:59';

SELECT * FROM #Graph ORDER BY YearNum, MonthNum, Center;


 
";

            var Grid = dbContext.Database.SqlQuery<Grid3>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName)).ToList();


            // Parse the user's date range once
            DateTime startDate = DateTime.Parse(req.StartDate);
            DateTime endDate = DateTime.Parse(req.EndDate);

            // Process with filtering, grouping by month, and removing duplicates
            var Grid3 = Grid
                .Select(x => new
                {
                    Record = x,
                    DateParsed = DateTime.TryParse(x.asDate, out var dt) ? dt : (DateTime?)null
                })
                .Where(x => x.DateParsed != null)
                .Where(x => x.DateParsed.Value >= startDate && x.DateParsed.Value <= endDate)
                .Where(x => x.Record.District.Contains(req.DistrictName) &&
                            x.Record.Center.Contains(req.CenterName))
                .GroupBy(x => new
                {
                    x.DateParsed.Value.Year,
                    x.DateParsed.Value.Month,
                    x.Record.Center
                })
                .Select(g => g.OrderByDescending(x => x.DateParsed).First().Record) // latest per center per month
                .ToList();


            response.Grid3 = Grid3;
        }
        private void MSUOpenClose(DashboardRequest req, DashboardResponse response)
        {
            string Query = $@"
 
-- MSU
WITH cte AS (  SELECT   s.ProjectID, sd.sbjnum,
        MAX(CASE WHEN sd.FieldId = 50482 THEN sd.[FieldValue] END) AS IsOpen,
        MAX(CASE WHEN sd.FieldId in( 50484) THEN sd.[FieldValue] END) AS District,
        MAX(CASE WHEN sd.FieldId = 50486 THEN sd.[FieldValue] END) AS Center
    FROM   SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum and Convert(datetime, s.DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY s.ProjectID, sd.sbjnum),
cte_with_titles AS (  SELECT cte.ProjectID,   cte.sbjnum,  cte.IsOpen, p.Title AS DistrictTitle,  pp.Title AS CenterTitle  FROM  cte
  INNER JOIN ProjectFieldSample p ON cte.District = p.Code AND p.FieldID = 50484 
    INNER JOIN ProjectFieldSample pp ON cte.Center = pp.Code AND pp.FieldID =  50486),
SplitStatus AS (    SELECT    cwt.ProjectID,   cwt.DistrictTitle,  cwt.CenterTitle,   value AS StatusCode
    FROM   cte_with_titles cwt   CROSS APPLY dbo.SplitStringValue(cwt.IsOpen, ',') 
 where (cwt.DistrictTitle like '%{req.DistrictName}%' or '' = '{req.DistrictName}') and (cwt.CenterTitle like  '%{req.CenterName}%' or '' = '{req.CenterName}')
)
SELECT 
    ProjectID, CASE    WHEN StatusCode = '1' THEN 'Open'  ELSE 'Close' END AS Title, COUNT(*) AS OpenClose
FROM   SplitStatus ss GROUP BY  ProjectID,  CASE   WHEN StatusCode = '1' THEN 'Open'  ELSE 'Close'   END ORDER BY  Title;
";


            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = "";
            }
            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = "";
            }
            var OpenClose = dbContext.Database.SqlQuery<PieChartOC>(Query);

            response.MSUOpenClose = OpenClose.ToList();
        }
        private void FWCOpenClose(DashboardRequest req, DashboardResponse response)
        {
            string Query = $@"
-- FWC

WITH cte AS (  SELECT   s.ProjectID, sd.sbjnum,
        MAX(CASE WHEN sd.FieldId = 55585 THEN sd.[FieldValue] END) AS IsOpen,
        MAX(CASE WHEN sd.FieldId in( 55587) THEN sd.[FieldValue] END) AS District,
        MAX(CASE WHEN sd.FieldId = 55588 THEN sd.[FieldValue] END) AS Center
    FROM   SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum and convert(datetime, s.DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY s.ProjectID, sd.sbjnum),
cte_with_titles AS (  SELECT cte.ProjectID,   cte.sbjnum,  cte.IsOpen, p.Title AS DistrictTitle,  pp.Title AS CenterTitle  FROM  cte
  INNER JOIN ProjectFieldSample p ON cte.District = p.Code AND p.FieldID = 55587 
    INNER JOIN ProjectFieldSample pp ON cte.Center = pp.Code AND pp.FieldID =  55588),
SplitStatus AS (    SELECT    cwt.ProjectID,   cwt.DistrictTitle,  cwt.CenterTitle,   value AS StatusCode
    FROM   cte_with_titles cwt   CROSS APPLY dbo.SplitStringValue(cwt.IsOpen, ',') 
 where (cwt.DistrictTitle like '%{req.DistrictName}%' or '' = '{req.DistrictName}') and (cwt.CenterTitle like  '%{req.CenterName}%' or '' = '{req.CenterName}')
)
SELECT 
    ProjectID, CASE    WHEN StatusCode = '1' THEN 'Open'  ELSE 'Close' END AS Title, COUNT(*) AS OpenClose
FROM   SplitStatus ss GROUP BY  ProjectID,  CASE   WHEN StatusCode = '1' THEN 'Open'  ELSE 'Close'   END ORDER BY  Title;";


            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = "";
            }
            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = "";
            }
            var OpenClose = dbContext.Database.SqlQuery<PieChartOC>(Query);

            response.FWCOpenClose = OpenClose.ToList();
        }
        private void RHSOpenClose(DashboardRequest req, DashboardResponse response)
        {
            string Query = $@"
-- RHS
WITH cte AS (  SELECT   s.ProjectID, sd.sbjnum,
        MAX(CASE WHEN sd.FieldId = 55570 THEN sd.[FieldValue] END) AS IsOpen,
        MAX(CASE WHEN sd.FieldId in( 50435) THEN sd.[FieldValue] END) AS District,
        MAX(CASE WHEN sd.FieldId = 50446 THEN sd.[FieldValue] END) AS Center
    FROM   SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum and convert(datetime, s.DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' GROUP BY s.ProjectID, sd.sbjnum),
cte_with_titles AS (  SELECT cte.ProjectID,   cte.sbjnum,  cte.IsOpen, p.Title AS DistrictTitle,  pp.Title AS CenterTitle  FROM  cte
  INNER JOIN ProjectFieldSample p ON cte.District = p.Code AND p.FieldID = 50435 
    INNER JOIN ProjectFieldSample pp ON cte.Center = pp.Code AND pp.FieldID =  50446),
SplitStatus AS (    SELECT    cwt.ProjectID,   cwt.DistrictTitle,  cwt.CenterTitle,   value AS StatusCode
    FROM   cte_with_titles cwt   CROSS APPLY dbo.SplitStringValue(cwt.IsOpen, ',') 
 where (cwt.DistrictTitle like '%{req.DistrictName}%' or '' = '{req.DistrictName}') and (cwt.CenterTitle like  '%{req.CenterName}%' or '' = '{req.CenterName}')
)
    
SELECT 
    ProjectID, CASE    WHEN StatusCode = '1' THEN 'Open'  ELSE 'Close' END AS Title, COUNT(*) AS OpenClose
FROM   SplitStatus ss GROUP BY  ProjectID,  CASE   WHEN StatusCode = '1' THEN 'Open'  ELSE 'Close'   END ORDER BY  Title;

";

            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = "";
            }
            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = "";
            }
            var OpenClose = dbContext.Database.SqlQuery<PieChartOC>(Query);

            response.RHSOpenClose = OpenClose.ToList();
        }
        public List<SDPsStatus> SDPStatus(DashboardRequest req)
        {
            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.DeviceTimestamp, 
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
	   
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
		sd6.fieldId as FieldId6, sd6.fieldValue as FieldValue6,
		sd7.fieldId as FieldId7, sd7.fieldValue as FieldValue7,
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.DeviceTimestamp desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (55591,50495,52571) -- Premises
	    inner join SurveyData sd6 on s.sbjnum = sd6.sbjnum and sd6.FieldId in (55570,50482,55585) -- Open Close Center Status
		inner join SurveyData sd7 on s.sbjnum = sd7.sbjnum and sd7.FieldId in (50437,50634,55590) -- Status 
		{Where})
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 
 as Premises,
case when FieldValue6 =1 then 'Open' else 'Close' end as OpenClose,
 FieldValue7 as Status, convert(varchar, DeviceTimestamp,101) asDate
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	inner join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)
	inner join ProjectFieldSample fs6 on cte.FieldId6 = fs6.FieldID and fs6.Code IN (cte.FieldValue6)
	inner join ProjectFieldSample fs7 on cte.FieldId7 = fs7.FieldID and fs7.Code IN (cte.FieldValue7)
    where RowNum = 1 and convert(datetime, DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' select * from #Graph";


            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }

            var con = dbContext.Database.SqlQuery<SDPsStatus>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));
            return con.ToList();
        }
        public List<StuffPosition> StuffDetailReportData(DashboardRequest req)
        {

            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	select  s.sbjnum, s.DeviceTimestamp, s.ProjectID,
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
	   
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
	
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.DeviceTimestamp desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (55592,50496,50635) {Where})
         	

select  convert(varchar, DeviceTimestamp,101) asDate,  fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 
 as Remarks
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	Left join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)

    where RowNum = 1 and len(FieldValue5)  between 1 and 19 and  convert(datetime,  DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' select * from #Graph

 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var stuffPosition = dbContext.Database.SqlQuery<StuffPosition>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));
            return stuffPosition.ToList();
        }
        public ContraceptiveStockPositionModel ContraceptiveStockPosition(DashboardRequest req)
        {
            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END;

;WITH cte AS (
    SELECT 
        s.ProjectID,  
        s.sbjnum, 
        s.DeviceTimestamp, 
        sd2.fieldId AS FieldId2, 
        sd2.fieldValue AS FieldValue2,
        sd3.fieldId AS FieldId3, 
        sd3.fieldValue AS FieldValue3,
        sd5.fieldId AS FieldId5, 
        sd5.fieldValue AS FieldValue5,
        ROW_NUMBER() OVER (
            PARTITION BY sd2.fieldId, sd2.fieldValue, sd3.fieldId, sd3.fieldValue, sd5.fieldId, sd5.fieldValue 
            ORDER BY s.DeviceTimestamp DESC
        ) AS RowNum
    FROM survey s
    INNER JOIN SurveyData sd2 
        ON s.sbjnum = sd2.sbjnum AND sd2.FieldId IN (50435, 50484, 55587) -- District
    INNER JOIN SurveyData sd3 
        ON s.sbjnum = sd3.sbjnum AND sd3.FieldId IN (50446, 50486, 55588) -- Center
    INNER JOIN SurveyData sd5 
        ON s.sbjnum = sd5.sbjnum AND sd5.FieldId IN (50559, 50504, 55601) -- ConStockPosition
    {Where}
), base AS (
    SELECT  
        (SELECT TOP 1 p.[Name] FROM Project p WHERE p.Id = ProjectID) AS ProjectName,
        fs2.Title AS District,
        fs3.Title AS Center, 
        ISNULL(FieldValue5,'') AS ConStockPosition,
        DeviceTimestamp,
        CONVERT(VARCHAR, DeviceTimestamp, 101) AS asDate,
        sbjnum
    FROM cte
    INNER JOIN ProjectFieldSample fs2 
        ON cte.FieldId2 = fs2.FieldID AND fs2.Code IN (cte.FieldValue2)
    INNER JOIN ProjectFieldSample fs3 
        ON cte.FieldId3 = fs3.FieldID AND fs3.Code IN (cte.FieldValue3)
    LEFT JOIN ProjectFieldSample fs5 
        ON cte.FieldId5 = fs5.FieldID AND fs5.Code IN (cte.FieldValue5)
    WHERE  RowNum = 1 AND LEN(FieldValue5) > 20 
      AND Convert(datetime, DeviceTimestamp,101) BETWEEN '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59'
), filtered AS (
    SELECT *,
        ROW_NUMBER() OVER (
            PARTITION BY Center, YEAR(DeviceTimestamp), MONTH(DeviceTimestamp)
            ORDER BY DeviceTimestamp ASC
        ) AS rn
    FROM base
)
SELECT 
    ProjectName,
    District,
    Center,
    ConStockPosition,
    asDate,
    sbjnum
INTO #Graph
FROM filtered
WHERE rn = 1;

-- Final result
SELECT * FROM #Graph order by asDate desc;

";

            var con = dbContext.Database.SqlQuery<Grid7>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName ?? "") && x.Center.Contains(req.CenterName ?? ""));
            List<ContraceptiveStockPositionResponse> dataTable = new List<ContraceptiveStockPositionResponse>();
            foreach (var item in con)
            {
                try
                {
                    var ConStockPosition = item.ConStockPosition.Split('|');
                    var Mon1 = ConStockPosition[0].Split('-')[1].Split(',')[0];
                    var Mon2 = ConStockPosition[1].Split('-')[1].Split(',')[0];
                    var Mon3 = ConStockPosition[2].Split('-')[1].Split(',')[0];
                    var Mon4 = ConStockPosition[3].Split('-')[1].Split(',')[0];
                    var Mon5 = ConStockPosition[4].Split('-')[1].Split(',')[0];
                    var Mon6 = ConStockPosition[5].Split('-')[1].Split(',')[0];
                    var Mon7 = ConStockPosition[6].Split('-')[1].Split(',')[0];
                    var Mon8 = ConStockPosition[7].Split('-')[1].Split(',')[0];
                    dataTable.Add(new ContraceptiveStockPositionResponse
                    {
                        Date = item.asDate,
                        SDP = item.ProjectName,
                        District = item.District,
                        Center = item.Center,
                        CondomsStock = int.TryParse(Mon1, out var v1) ? v1 : 0,
                        POP = int.TryParse(Mon2, out var v2) ? v2 : 0,
                        COC = int.TryParse(Mon3, out var v3) ? v3 : 0,
                        ECP = int.TryParse(Mon4, out var v4) ? v4 : 0,
                        ThreemonthsInj = int.TryParse(Mon5, out var v5) ? v5 : 0,
                        DefoStock = int.TryParse(Mon6, out var v6) ? v6 : 0,
                        IUD = int.TryParse(Mon7, out var v7) ? v7 : 0,
                        Jadelle = int.TryParse(Mon8, out var v8) ? v8 : 0,
                    });
                }
                catch (Exception)
                {

                }
            }


            var groupedData = dataTable
                .GroupBy(item => item.SDP)
                .Select(g => new ContraceptiveStockPositionStockWise
                {
                    SDP = g.Key,
                    CondomsStock = g.Sum(x => x.CondomsStock),
                    POP = g.Sum(x => x.POP),
                    COC = g.Sum(x => x.COC),
                    ECP = g.Sum(x => x.ECP),
                    ThreemonthsInj = g.Sum(x => x.ThreemonthsInj),
                    DefoStock = g.Sum(x => x.DefoStock),
                    IUD = g.Sum(x => x.IUD),
                    Jadelle = g.Sum(x => x.Jadelle)
                })
                .ToList();


            var alerts = dataTable
                .GroupBy(x => new { x.Center, x.District })
                .Select(g =>
                {
                    int totalStock = g.Sum(x =>
                        x.CondomsStock + x.POP + x.COC + x.ECP +
                        x.ThreemonthsInj + x.DefoStock + x.IUD + x.Jadelle);

                    string level, message;

                    if (totalStock < 5)
                    {
                        level = "low";
                        message = "🔴 Very low contraceptive stock";
                    }
                    else if (totalStock < 1000)
                    {
                        level = "medium";
                        message = "🟠 Medium stock levels";
                    }
                    else if (totalStock < 10000)
                    {
                        level = "perfect";
                        message = "🟡 Good stock levels";
                    }
                    else
                    {
                        level = "full";
                        message = "🟢 Full stock available";
                    }

                    return new CenterStockAlert
                    {
                        Center = g.Key.Center,
                        District = g.Key.District,
                        AlertLevel = level,
                        Message = message,
                        TotalStock = totalStock
                    };
                })
                .Where(x => x.AlertLevel != "full") // Skip full-stock if not needed
                .OrderBy(x => x.TotalStock)         // ✅ Show lowest stock first
                .ToList();


            ContraceptiveStockPositionModel response = new ContraceptiveStockPositionModel();
            response.contraceptiveStockPositionResponse = dataTable;
            response.contraceptiveStockPositionStockWise = groupedData;
            response.centerStockAlerts = alerts;

            return response;
        }
        public List<Grid14> FuniturePosition(DashboardRequest req)
        {

            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.DeviceTimestamp, 
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
	    sd6.fieldId as FieldId6, sd6.fieldValue as FieldValue6,

	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.DeviceTimestamp desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (50512,50474,55615) -- FP
		Inner join SurveyData sd6 on s.sbjnum = sd6.sbjnum and sd6.FieldId in (50513,53570,55616) -- FPQ
		{Where}
		)
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 as FP,
  FieldValue6 as FPQ,
 cte.FieldId5,

  convert(varchar, DeviceTimestamp,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	inner join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID 
	inner join ProjectFieldSample fs6 on cte.FieldId6 = fs6.FieldID 
	
    where RowNum = 1 and len(FieldValue6) > 5 and DeviceTimestamp  BETWEEN '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' select distinct * from #Graph   

 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var con = dbContext.Database.SqlQuery<Grid14>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));
            return con.ToList();
        }
        public List<StockOfContraceptiveResponse> StockOfConteraceptives(DashboardRequest req)
        {

            string Sql = $@"select sd.fieldValue as StockOfContraceptive, sd.sbjnum,ProjectId,convert(varchar, DeviceTimestamp,101)   as asDate
	                        from SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum 
                             where sd.FieldId in(55601,50504,50559) and s.ProjectID in ({req.ProjectId}) and 
convert(datetime, DeviceTimestamp,101) between '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' ";


            var con = dbContext.Database.SqlQuery<StockOfContraceptiveResponse>(Sql).ToList();


            List<Comodities> stock = ParseComoditiesFromResponses(con);



            return con;
        }
        public static List<Comodities> ParseComoditiesFromResponses(List<StockOfContraceptiveResponse> responses)
        {
            var allStocks = new List<Comodities>();

            foreach (var response in responses)
            {
                if (string.IsNullOrWhiteSpace(response.StockOfContraceptive))
                    continue;

                var parts = response.StockOfContraceptive
                    .Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

                var comodity = new Comodities();
                string commonDate = null;

                foreach (var part in parts)
                {
                    var split = part.Split(new[] { '-' }, 2);
                    if (split.Length < 2) continue;

                    string key = split[0].Trim();
                    string remainder = split[1].Trim();

                    string qty = "";
                    string date = "";

                    if (remainder.Contains(","))
                    {
                        var vals = remainder.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                        DateTime dummy;
                        if (DateTime.TryParse(vals[0], out dummy))
                        {
                            date = vals[0];
                            qty = vals.Length > 1 ? vals[1] : "";
                        }
                        else
                        {
                            qty = vals[0];
                            date = vals.Length > 1 ? vals[1] : "";
                        }
                    }

                    if (commonDate == null)
                        commonDate = date;

                    switch (key)
                    {
                        case "1": comodity.Condoms = qty; break;
                        case "2": comodity.POP = qty; break;
                        case "3": comodity.COC = qty; break;
                        case "4": comodity.ECP = qty; break;
                        case "5": comodity.ThreeMonth = qty; break;
                        case "6": comodity.Defo = qty; break;
                        case "7": comodity.IUD = qty; break;
                        case "8": comodity.Jodelle = qty; break;
                    }
                }

                comodity.Date = commonDate ?? response.asDate; // Make sure this property exists
                allStocks.Add(comodity);
            }

            return allStocks;
        }
        public List<PdfDetailReport> MonitoringVisitsReport(DashboardRequest req, string name, IPrincipal User)
        {
            try
            {
                //var District = string.Empty;
                //bool isAdmin = false;
                //if (User.IsInRole("Admin"))
                //{
                //    var GetDistrict = name.Split('_');
                //    var removeAtRat = GetDistrict[1].Split('@');
                //    District = removeAtRat[0];
                //    isAdmin = true;
                //}
                //else
                //{
                //    if (name.Contains("_"))
                //    {
                //        isAdmin = false;
                //    }
                //    else
                //    {
                //        District = "";
                //    }
                //}

                string Query = $@"
IF OBJECT_ID('tempdb..#SurveyReport') IS NOT NULL
    DROP TABLE #SurveyReport

select Convert(varchar,s.Longitude) Longitude, Convert(varchar,s.Latitude) Latitude ,s.sbjnum, Convert(varchar,s.Created,101) Created, s.SurveyorName,Convert(varchar,s.DeviceTimestamp,101) as DeviceTimestamp,
Convert(varchar,isnull((select top 1 sd.FieldValue from SurveyData sd where sd.FieldId in (50435,50484,55587) and sd.sbjnum = s.sbjnum),0)) as District,
Convert(varchar,isnull((select top 1 sd.FieldValue from SurveyData sd where sd.FieldId in (50446,50486,55588) and sd.sbjnum = s.sbjnum),0)) as Center,
Convert(varchar,isnull((select top 1 sd.FieldId from SurveyData sd where sd.FieldId in (50435,50484,55587) and sd.sbjnum = s.sbjnum),0)) as DistrictFieldID,
Convert(varchar,isnull((select top 1  sd.FieldId from SurveyData sd where sd.FieldId in (50446,50486,55588) and sd.sbjnum = s.sbjnum),0)) as CenterFieldId
, case when s.projectID = 7120 then 'RHS-S' when  s.projectID = 7121 then 'MSU' when s.projectID = 7122 then 'FWC' else '' end as Project
into #SurveyReport
from Survey  s 
where s.projectID in ({req.ProjectId}) and s.DeviceTimestamp  BETWEEN '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59' order by s.sbjnum desc
 select sp.*,isnull(pfD.Title,'') DistrictName, isnull(pfC.Title,'') CenterName  from #SurveyReport sp 
 Left join   ProjectFieldSample pfC on sp.Center = pfC.Code and sp.CenterFieldId = pfC.FieldID
 Left join ProjectFieldSample pfD on sp.District = pfD.Code and sp.DistrictFieldID = pfD.FieldID 
 order by DeviceTimestamp desc
";

                if (string.IsNullOrEmpty(req.DistrictName))
                {
                    req.DistrictName = string.Empty;
                }

                if (string.IsNullOrEmpty(req.CenterName))
                {
                    req.CenterName = string.Empty;
                }
                var GetSurvey = dbContext.Database.SqlQuery<PdfDetailReport>(Query);
                if (true) // isAdmin
                {
                    return GetSurvey.ToList();
                }
                else
                {
                    var DistrictWise = GetSurvey.Where(x => x.DistrictName.ToUpper().Contains(req.DistrictName) && x.CenterName.Contains(req.CenterName)).ToList();
                    return DistrictWise.ToList();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public string GetReport(string sbjnum, string Heading, string BaseUrl)
        {
            string url = string.Empty;
            if (Heading.ToUpper() == "FWC")
            {
                url = $@"{BaseUrl}/Designer/FWCReport?id={sbjnum}&Heading=FWC";
            }
            else if (Heading.ToUpper().Contains("RHS"))
            {
                url = $@"{BaseUrl}/Designer/RHSReport?id={sbjnum}&Heading=RHS-S";
            }
            else
            {
                url = $@"{BaseUrl}/Designer/MSUReport?id={sbjnum}&Heading=UMS";
            }
            return url;
        }
        public List<SurveyReportViewModel> OfficerVisitReport()
        {
            try
            {

                string EndDate = DateTime.Now.ToString("MM/dd/yyyy");
                string StartDateDat = DateTime.Now.AddDays(-7).ToString("MM/dd/yyyy");


                string Query = $@"
IF OBJECT_ID('tempdb..#SurveyReport') IS NOT NULL
    DROP TABLE #SurveyReport;

SELECT 
    CONVERT(VARCHAR, ISNULL(
        (SELECT TOP 1 sd.FieldValue FROM SurveyData sd WHERE sd.FieldId IN (50435,50484,55587) AND sd.sbjnum = s.sbjnum), 0)) AS District,
    s.SurveyorName AS OfficerName, 
    s.SurveyorName AS DesignationName,
    s.DeviceTimestamp,
    CONVERT(VARCHAR, ISNULL(
        (SELECT TOP 1 sd.FieldValue FROM SurveyData sd WHERE sd.FieldId IN (50446,50486,55588) AND sd.sbjnum = s.sbjnum), 0)) AS Center,
    CONVERT(VARCHAR, ISNULL(
        (SELECT TOP 1 sd.FieldId FROM SurveyData sd WHERE sd.FieldId IN (50435,50484,55587) AND sd.sbjnum = s.sbjnum), 0)) AS DistrictFieldID,
    CONVERT(VARCHAR, ISNULL(
        (SELECT TOP 1 sd.FieldId FROM SurveyData sd WHERE sd.FieldId IN (50446,50486,55588) AND sd.sbjnum = s.sbjnum), 0)) AS CenterFieldId,
    CASE 
        WHEN s.projectID = 7120 THEN 'RHS-S' 
        WHEN s.projectID = 7121 THEN 'MSU' 
        WHEN s.projectID = 7122 THEN 'FWC' 
        ELSE '' 
    END AS Project
INTO #SurveyReport
FROM Survey s 
WHERE s.projectID IN (7120,7121,7122) 
    AND s.DeviceTimestamp BETWEEN '2025-10-16 00:00:01' AND '2025-10-23 12:59:59'
ORDER BY s.sbjnum DESC;

WITH DistinctVisits AS (
    SELECT 
        ISNULL(pfD.Title, '') AS District,
        sp.OfficerName,
        sp.DesignationName,
        sp.Project,
        CONVERT(DATE, sp.DeviceTimestamp) AS VisitDate, -- date only
        sp.Center
    FROM #SurveyReport sp
    LEFT JOIN ProjectFieldSample pfD ON sp.District = pfD.Code AND sp.DistrictFieldID = pfD.FieldID
)

SELECT 
    District,
    OfficerName,
    DesignationName,

    -- RHS-S total centers / visited centers
    CAST(
        (SELECT COUNT(*) FROM ProjectFieldSample WHERE Title LIKE '%' + District + '%' AND Title LIKE '%RHS%')
        AS VARCHAR(10)) + '/' +
    CAST(
        COUNT(DISTINCT CASE WHEN Project = 'RHS-S' THEN CAST(Center AS VARCHAR) + '|' + CAST(VisitDate AS VARCHAR) END)
        AS VARCHAR(10)) AS [RHSCountVisited],

    -- FWC total centers / visited centers
    CAST(
        (SELECT COUNT(*) FROM ProjectFieldSample WHERE Title LIKE '%' + District + '%' AND Title LIKE '%FWC%')
        AS VARCHAR(10)) + '/' +
    CAST(
        COUNT(DISTINCT CASE WHEN Project = 'FWC' THEN CAST(Center AS VARCHAR) + '|' + CAST(VisitDate AS VARCHAR) END)
        AS VARCHAR(10)) AS [FWCCountVisited],

    -- MSU total centers / visited centers
    CAST(
        (SELECT COUNT(*) FROM ProjectFieldSample WHERE Title LIKE '%' + District + '%' AND Title LIKE '%MSU%')
        AS VARCHAR(10)) + '/' +
    CAST(
        COUNT(DISTINCT CASE WHEN Project = 'MSU' THEN CAST(Center AS VARCHAR) + '|' + CAST(VisitDate AS VARCHAR) END)
        AS VARCHAR(10)) AS [MSUCountVisited],

    -- Total distinct centers visited overall (any project)

	  CAST(
    (SELECT COUNT(*) FROM ProjectFieldSample WHERE Title LIKE '%' + District + '%'     AND (Title LIKE '%FWC%' or   Title LIKE '%RHS%' or   Title LIKE '%MSU%' )) 
    AS VARCHAR(10)
) + '/' + 
CAST(
    COUNT(DISTINCT CAST(Center AS VARCHAR) + '|' + CAST(VisitDate AS VARCHAR))
    AS VARCHAR(10)
) AS TotalDistinctCentersVisited


FROM DistinctVisits
GROUP BY 
    District,
    OfficerName,
    DesignationName
ORDER BY
    District;

";


                var GetSurvey = dbContext.Database.SqlQuery<SurveyReportViewModel>(Query);
                return GetSurvey.ToList();

            }
            catch (Exception)
            {
                throw;
            }
        }
        public List<ContraceptiveStock> DetailOfContraceptive(DashboardRequest req)
        {


            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.Created, 
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
	   
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
	

	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (50559,50504,55601) -- ConStockPosition
       where s.ProjectID in ({req.ProjectId})
		
		)
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs2.Title as District ,fs3.Title as Center, 
 isnull(FieldValue5,'')  
 as ConStockPosition,


  convert(varchar, Created,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	left join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)

	
    where RowNum = 1 and len(FieldValue5) > 20 and Convert(datetime, Created,101) between '{req.StartDate}' and '{req.EndDate}' select * from #Graph   

 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var Stock = dbContext.Database.SqlQuery<Grid7>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));


            List<ContraceptiveStock> StockList = new List<ContraceptiveStock>();

            foreach (var item in Stock)
            {
                try
                {

                    var ConStockPosition = item.ConStockPosition.Split('|');

                    var Condoms = ConStockPosition[0].ToString().Split('-')[1].Split(',')[0];
                    var CondomsExpiry = ConStockPosition[0].ToString().Split(',')[1];

                    var POP = ConStockPosition[1].ToString().Split('-')[1].Split(',')[0];
                    var POPExpiry = ConStockPosition[1].ToString().Split(',')[1];

                    var COC = ConStockPosition[2].ToString().Split('-')[1].Split(',')[0];
                    var COCExpiry = ConStockPosition[2].ToString().Split(',')[1];

                    var ECP = ConStockPosition[3].ToString().Split('-')[1].Split(',')[0];
                    var ECPExpiry = ConStockPosition[3].ToString().Split(',')[1];

                    var ThreeMonth = ConStockPosition[4].ToString().Split('-')[1].Split(',')[0];
                    var ThreeMonthExpiry = ConStockPosition[4].ToString().Split(',')[1];

                    var Defo = ConStockPosition[5].ToString().Split('-')[1].Split(',')[0];
                    var DefoExpiry = ConStockPosition[5].ToString().Split(',')[1];

                    var IUD = ConStockPosition[6].ToString().Split('-')[1].Split(',')[0];
                    var IUDExpiry = ConStockPosition[6].ToString().Split(',')[1];

                    var Jadelle = ConStockPosition[7].ToString().Split('-')[1].Split(',')[0];
                    var JadelleExpiry = ConStockPosition[7].ToString().Split(',')[1];

                    StockList.Add(new ContraceptiveStock
                    {
                        SDP = item.ProjectName,
                        District = item.District,
                        Center = item.Center,
                        Date = item.asDate,
                        Condoms = Condoms + "|" + CondomsExpiry,
                        POP = POP + "|" + POPExpiry,
                        COC = COC + "|" + COCExpiry,
                        Defo = Defo + "|" + DefoExpiry,
                        ECP = ECP + "|" + ECPExpiry,
                        ThreeMonth = ThreeMonth + "|" + ThreeMonthExpiry,
                        IUD = IUD + "|" + IUDExpiry,
                        Jadelle = Jadelle + "|" + JadelleExpiry,
                    });

                }
                catch (Exception)
                {

                }
            }
            return StockList;
        }
        public List<EquiptmentModel> EquiptmentPosition(DashboardRequest req)
        {

            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END;

;WITH cte AS (
    SELECT 
        s.ProjectID,  
        s.sbjnum, 
        s.DeviceTimestamp, 
        sd2.fieldId AS FieldId2, 
        sd2.fieldValue AS FieldValue2,
        sd3.fieldId AS FieldId3, 
        sd3.fieldValue AS FieldValue3,
        sd5.fieldId AS FieldId5, 
        sd5.fieldValue AS FieldValue5,
        ROW_NUMBER() OVER (
            PARTITION BY sd2.fieldId, sd2.fieldValue, sd3.fieldId, sd3.fieldValue, sd5.fieldId, sd5.fieldValue 
            ORDER BY s.DeviceTimestamp DESC
        ) AS RowNum
    FROM survey s
    INNER JOIN SurveyData sd2 
        ON s.sbjnum = sd2.sbjnum AND sd2.FieldId IN (50435, 50484, 55587) -- District
    INNER JOIN SurveyData sd3 
        ON s.sbjnum = sd3.sbjnum AND sd3.FieldId IN (50446, 50486, 55588) -- Center
    INNER JOIN SurveyData sd5 
        ON s.sbjnum = sd5.sbjnum AND sd5.FieldId IN (55613, 50510, 50471) -- EuiptmentPosition
    {Where}
), base AS (
    SELECT  
        (SELECT TOP 1 p.[Name] FROM Project p WHERE p.Id = ProjectID) AS ProjectName,
        fs2.Title AS District,
        fs3.Title AS Center, 
        ISNULL(FieldValue5,'') AS EuiptmentPosition,
        DeviceTimestamp,
        CONVERT(VARCHAR, DeviceTimestamp, 101) AS asDate,
        sbjnum
    FROM cte
    INNER JOIN ProjectFieldSample fs2 
        ON cte.FieldId2 = fs2.FieldID AND fs2.Code IN (cte.FieldValue2)
    INNER JOIN ProjectFieldSample fs3 
        ON cte.FieldId3 = fs3.FieldID AND fs3.Code IN (cte.FieldValue3)
    LEFT JOIN ProjectFieldSample fs5 
        ON cte.FieldId5 = fs5.FieldID AND fs5.Code IN (cte.FieldValue5)
    WHERE  RowNum = 1 AND LEN(FieldValue5) > 20 
      AND Convert(datetime, DeviceTimestamp,101) BETWEEN '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59'
), filtered AS (
    SELECT *,
        ROW_NUMBER() OVER (
            PARTITION BY Center, YEAR(DeviceTimestamp), MONTH(DeviceTimestamp)
            ORDER BY DeviceTimestamp ASC
        ) AS rn
    FROM base
)
SELECT 
    ProjectName,
    District,
    Center,
    EuiptmentPosition,
    asDate,
    sbjnum
INTO #Graph
FROM filtered
WHERE rn = 1;

-- Final result
SELECT * FROM #Graph order by asDate desc;
 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var con = dbContext.Database.SqlQuery<EquiptmentModel>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));
            return con.ToList();
        }
        public EquiptmentPositionResponse EquiptmentPositionStock(DashboardRequest req)
        {

            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END;

;WITH cte AS (
    SELECT 
        s.ProjectID,  
        s.sbjnum, 
        s.DeviceTimestamp, 
        sd2.fieldId AS FieldId2, 
        sd2.fieldValue AS FieldValue2,
        sd3.fieldId AS FieldId3, 
        sd3.fieldValue AS FieldValue3,
        sd5.fieldId AS FieldId5, 
        sd5.fieldValue AS FieldValue5,
        ROW_NUMBER() OVER (
            PARTITION BY sd2.fieldId, sd2.fieldValue, sd3.fieldId, sd3.fieldValue, sd5.fieldId, sd5.fieldValue 
            ORDER BY s.DeviceTimestamp DESC
        ) AS RowNum
    FROM survey s
    INNER JOIN SurveyData sd2 
        ON s.sbjnum = sd2.sbjnum AND sd2.FieldId IN (50435, 50484, 55587) -- District
    INNER JOIN SurveyData sd3 
        ON s.sbjnum = sd3.sbjnum AND sd3.FieldId IN (50446, 50486, 55588) -- Center
    INNER JOIN SurveyData sd5 
        ON s.sbjnum = sd5.sbjnum AND sd5.FieldId IN (55613, 50510, 50471) -- EuiptmentPosition
    {Where}
), base AS (
    SELECT  
        (SELECT TOP 1 p.[Name] FROM Project p WHERE p.Id = ProjectID) AS ProjectName,
        fs2.Title AS District,
        fs3.Title AS Center, 
        ISNULL(FieldValue5,'') AS EuiptmentPosition,
        DeviceTimestamp,
        CONVERT(VARCHAR, DeviceTimestamp, 101) AS asDate,
        sbjnum
    FROM cte
    INNER JOIN ProjectFieldSample fs2 
        ON cte.FieldId2 = fs2.FieldID AND fs2.Code IN (cte.FieldValue2)
    INNER JOIN ProjectFieldSample fs3 
        ON cte.FieldId3 = fs3.FieldID AND fs3.Code IN (cte.FieldValue3)
    LEFT JOIN ProjectFieldSample fs5 
        ON cte.FieldId5 = fs5.FieldID AND fs5.Code IN (cte.FieldValue5)
    WHERE  RowNum = 1 AND LEN(FieldValue5) > 20 
      AND Convert(datetime, DeviceTimestamp,101) BETWEEN '{req.StartDate} 00:00:01' and '{req.EndDate} 12:59:59'
), filtered AS (
    SELECT *,
        ROW_NUMBER() OVER (
            PARTITION BY Center, YEAR(DeviceTimestamp), MONTH(DeviceTimestamp)
            ORDER BY DeviceTimestamp ASC
        ) AS rn
    FROM base
)
SELECT 
    ProjectName,
    District,
    Center,
    EuiptmentPosition,
    asDate,
    sbjnum
INTO #Graph
FROM filtered
WHERE rn = 1;

-- Final result
SELECT * FROM #Graph order by asDate desc;
 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var con = dbContext.Database.SqlQuery<EquiptmentModel>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));


            // Create the response
            EquiptmentPositionResponse calculate = new EquiptmentPositionResponse();
            foreach (var item in con)
            {

                var data = item.EuiptmentPosition.Split(',');
                try
                {
                    // Convert all values to int, safely
                    var numbers = data.Select(x => int.TryParse(x, out var n) ? n : 0).ToArray();

         

                    calculate.Minilapkits += numbers.Length > 0 ? (numbers[0] == 2 || numbers[0] == 3 ? 1 : 0) : 0;
                    calculate.Iudkits += numbers.Length > 1 ? (numbers[1] == 2 || numbers[1] == 3 ? 1 : 0) : 0;
                    calculate.BPApparatus += numbers.Length > 2 ? (numbers[2] == 2 || numbers[2] == 3 ? 1 : 0) : 0;
                    calculate.Stethoscope += numbers.Length > 3 ? (numbers[3] == 2 || numbers[3] == 3 ? 1 : 0) : 0;
                    calculate.Thermometer += numbers.Length > 4 ? (numbers[4] == 2 || numbers[4] == 3 ? 1 : 0) : 0;
                    calculate.WeightingMachine += numbers.Length > 5 ? (numbers[5] == 2 || numbers[5] == 3 ? 1 : 0) : 0;
                    calculate.Stove += numbers.Length > 6 ? (numbers[6] == 2 || numbers[6] == 3 ? 1 : 0) : 0;
                    calculate.OTLights += numbers.Length > 7 ? (numbers[7] == 2 || numbers[7] == 3 ? 1 : 0) : 0;
                    calculate.HydrolicTable += numbers.Length > 8 ? (numbers[8] == 2 || numbers[8] == 3 ? 1 : 0) : 0;
                    calculate.Autoclave += numbers.Length > 9 ? (numbers[9] == 2 || numbers[9] == 3 ? 1 : 0) : 0;
                    calculate.OxygenCylinder += numbers.Length > 10 ? (numbers[10] == 2 || numbers[10] == 3 ? 1 : 0) : 0;
                    calculate.AspiratingPumps += numbers.Length > 11 ? (numbers[11] == 2 || numbers[11] == 3 ? 1 : 0) : 0;
                    calculate.WheelChair += numbers.Length > 12 ? (numbers[12] == 2 || numbers[12] == 3 ? 1 : 0) : 0;
                    calculate.Stretcher += numbers.Length > 13 ? (numbers[13] == 2 || numbers[13] == 3 ? 1 : 0) : 0;
                    calculate.Generators += numbers.Length > 14 ? (numbers[14] == 2 || numbers[14] == 3 ? 1 : 0) : 0;
                    calculate.Screen += numbers.Length > 15 ? (numbers[15] == 2 || numbers[15] == 3 ? 1 : 0) : 0;


                    calculate.MinilapkitsTotal += 1;
                    calculate.IudkitsTotal += 1;
                    calculate.BPApparatusTotal += 1;
                    calculate.StethoscopeTotal += 1;
                    calculate.ThermometerTotal += 1;
                    calculate.WeightingMachineTotal +=1;
                    calculate.StoveTotal += 1;
                    calculate.OTLightsTotal += 1;
                    calculate.HydrolicTableTotal += 1;
                    calculate.AutoclaveTotal +=1;
                    calculate.OxygenCylinderTotal += 1;
                    calculate.AspiratingPumpsTotal += 1;
                    calculate.WheelChairTotal +=1;
                    calculate.StretcherTotal += 1;
                    calculate.GeneratorsTotal +=1;
                    calculate.ScreenTotal += 1;
                }
                catch (Exception ex)
                {
                    // Handle exception (logging or ignore)
                }
 

            }
            return calculate;
        }

        public dynamic TechnicalMonitoringDetail(DashboardRequest req)
        {
            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.Created, 
       
		sd1.fieldId as FieldId1, sd1.fieldValue as FieldValue1,
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,

	
	row_number() over (partition by  sd1.fieldId, sd1.fieldValue,sd2.fieldId,sd2.fieldValue ,sd3.fieldId,sd3.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd1 on s.sbjnum = sd1.sbjnum and sd1.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50475,55573,55619) -- Technical
        {Where}
		
		)
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs1.Title as District ,fs2.Title as Center, 
 FieldValue3 
 as Technical,


  convert(varchar, Created,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs1 on cte.FieldId1 = fs1.FieldID and fs1.Code IN (cte.FieldValue1)
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID --and fs3.Code IN (cte.FieldValue3)  
  and created between '{req.StartDate}' and '{req.EndDate}' select distinct * from #Graph g where len(g.Technical) > 1

 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var con = dbContext.Database.SqlQuery<Grid11>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));
            return con;
        }

        public dynamic TechnicalMonitoringStock(DashboardRequest req)
        {
            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.Created, 
       
		sd1.fieldId as FieldId1, sd1.fieldValue as FieldValue1,
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,

	
	row_number() over (partition by  sd1.fieldId, sd1.fieldValue,sd2.fieldId,sd2.fieldValue ,sd3.fieldId,sd3.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd1 on s.sbjnum = sd1.sbjnum and sd1.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50475,55573,55619) -- Technical
        {Where}
		
		)
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs1.Title as District ,fs2.Title as Center, 
 FieldValue3 
 as Technical,


  convert(varchar, Created,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs1 on cte.FieldId1 = fs1.FieldID and fs1.Code IN (cte.FieldValue1)
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID --and fs3.Code IN (cte.FieldValue3)  
  and created between '{req.StartDate}' and '{req.EndDate}' select distinct * from #Graph g where len(g.Technical) > 1

 
";
            if (string.IsNullOrEmpty(req.DistrictName))
            {
                req.DistrictName = string.Empty;
            }

            if (string.IsNullOrEmpty(req.CenterName))
            {
                req.CenterName = string.Empty;
            }
            var con = dbContext.Database.SqlQuery<Grid11>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));


            MonitoringResponse monitoringResponse = new MonitoringResponse();

            foreach (var item in con)
            {
                var data = item.Technical.Split(',');
                monitoringResponse.HandWashing += data[0] == "1" ? 1 : 0;
                monitoringResponse.Decontamination += data[1] == "1" ? 1 : 0;
                monitoringResponse.Cleaning += data[2] == "1" ? 1 : 0;
                monitoringResponse.disinfection += data[3] == "1" ? 1 : 0;
                monitoringResponse.Wastedisposal += data[4] == "1" ? 1 : 0;

                monitoringResponse.HandWashingNo += data[0] == "2" ? 1 : 0;
                monitoringResponse.DecontaminationNo += data[1] == "2" ? 1 : 0;
                monitoringResponse.CleaningNo += data[2] == "2" ? 1 : 0;
                monitoringResponse.disinfectionNo += data[3] == "2" ? 1 : 0;
                monitoringResponse.WastedisposalNo += data[4] == "2" ? 1 : 0;
            }


            return monitoringResponse;
        }


        public List<IECMatrialResponse> IECMatrialStock(DashboardRequest req)
        {


            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.Created, 
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
		sd6.fieldId as FieldId6, sd6.fieldValue as FieldValue6,
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (50563,50614,55605) -- IECMatrial
	    Inner join SurveyData sd6 on s.sbjnum = sd6.sbjnum and sd6.FieldId in (52570,50615,55606) -- MECWheel
		{Where})
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 
 as IECMatrial,
FieldValue6 as MECWheel,convert(varchar, Created,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	left join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)
	left join ProjectFieldSample fs6 on cte.FieldId6 = fs6.FieldID and fs6.Code IN (cte.FieldValue6)
    where RowNum = 1 and  len(FieldValue5) > 4 and created between '{req.StartDate}' and '{req.EndDate}' select * from #Graph ";
            var IECMatrial = dbContext.Database.SqlQuery<Grid6>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));

            List<IECMatrialResponse> iECMatrialResponse = new List<IECMatrialResponse>();

            foreach (var item in IECMatrial)
            {

            }


            return iECMatrialResponse;
        }

        public List<Grid6> IECMatrialDetail(DashboardRequest req)
        {


            string Where = $"where s.ProjectID in ({req.ProjectId})";
            string Sql = $@"IF OBJECT_ID('tempdb..#Graph') IS NOT NULL
BEGIN
    DROP TABLE #Graph;
END

;with cte as (
	  select s.ProjectID,  s.sbjnum, s.Created, 
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
		sd6.fieldId as FieldId6, sd6.fieldValue as FieldValue6,
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (50563,50614,55605) -- IECMatrial
	    Inner join SurveyData sd6 on s.sbjnum = sd6.sbjnum and sd6.FieldId in (52570,50615,55606) -- MECWheel
		{Where})
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 
 as IECMatrial,
FieldValue6 as MECWheel,convert(varchar, Created,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	left join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)
	left join ProjectFieldSample fs6 on cte.FieldId6 = fs6.FieldID and fs6.Code IN (cte.FieldValue6)
    where RowNum = 1 and  len(FieldValue5) > 4 and created between '{req.StartDate}' and '{req.EndDate}' select * from #Graph ";
            var IECMatrial = dbContext.Database.SqlQuery<Grid6>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName));
            return IECMatrial.ToList();
        }
    }
}