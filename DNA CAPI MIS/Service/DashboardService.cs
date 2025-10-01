using DNA_CAPI_MIS.Models;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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
            var All = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID in( 7120,7121,7122) and Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";
            var RHS = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID = 7120 and Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";
            var MSU = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID = 7121 and Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";
            var FWC = $@"SELECT SurveyorName, COUNT(*) AS SurveyCount FROM Survey WHERE ProjectID = 7122 and Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY SurveyorName ORDER BY COUNT(*) DESC";

            var RHS_MSU_FWC = dbContext.Database.SqlQuery<SurveyorStats>(All);
            var queryRHS = dbContext.Database.SqlQuery<SurveyorStats>(RHS);
            var queryMSU = dbContext.Database.SqlQuery<SurveyorStats>(MSU);
            var queryFWC = dbContext.Database.SqlQuery<SurveyorStats>(FWC);

            if (RHS_MSU_FWC.Count() > 0)
            {
                response.All.SurveyCount = RHS_MSU_FWC.Sum(x => x.SurveyCount);
            }
            if (queryRHS.Count() > 0)
            {
                response.RHS.SurveyCount = queryRHS.Sum(x => x.SurveyCount);
            }
            if (queryMSU.Count() > 0)
            {
                response.MSU.SurveyCount = queryMSU.Sum(x => x.SurveyCount);
            }
            if (queryFWC.Count() > 0)
            {
                response.FWC.SurveyCount = queryFWC.Sum(x => x.SurveyCount);
            }
        }

        public void StatusOfBuilding(DashboardRequest req, DashboardResponse response)
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
		sd7.fieldId as FieldId7, sd7.fieldValue as FieldValue7,
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue,sd7.fieldValue,sd7.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (55594,50498,50461) -- indicate Sign
	    Inner join SurveyData sd6 on s.sbjnum = sd6.sbjnum and sd6.FieldId in (50557,50500,55595) -- Status of Building
		left  join SurveyData sd7 on s.sbjnum = sd7.sbjnum and sd7.FieldId in (50462,50499,55596) -- Cleanliess
		{Where} )
select  (select top 1 p.[Name] from Project p where p.Id=  ProjectID) as ProjectName, fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 
 as IndicateSign,
FieldValue6 as StatusOfBuilding,
 case when  FieldValue7 = '1' then 'Satisfactory' when FieldValue7 = '2' then 'Not Satisfactory' else '' end  as Cleanliness ,
  convert(varchar, Created,101) asDate,
 sbjnum
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	left join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)
	left join ProjectFieldSample fs6 on cte.FieldId6 = fs6.FieldID and fs6.Code IN (cte.FieldValue6)
	left join ProjectFieldSample fs7 on cte.FieldId7 = fs7.FieldID and fs7.Code IN (cte.FieldValue7)
    where RowNum = 1 and created between '{req.StartDate}' and '{req.EndDate}' select * from #Graph   

 
";

            var Grid = dbContext.Database.SqlQuery<Grid3>(Sql).ToList().Where(x => x.District.Contains(req.DistrictName) && x.Center.Contains(req.CenterName)).ToList();
            response.Grid3 = Grid;
        }

        private void MSUOpenClose(DashboardRequest req, DashboardResponse response)
        {
            string Query = $@"
 
-- MSU
WITH cte AS (  SELECT   s.ProjectID, sd.sbjnum,
        MAX(CASE WHEN sd.FieldId = 50482 THEN sd.[FieldValue] END) AS IsOpen,
        MAX(CASE WHEN sd.FieldId in( 50484) THEN sd.[FieldValue] END) AS District,
        MAX(CASE WHEN sd.FieldId = 50486 THEN sd.[FieldValue] END) AS Center
    FROM   SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum and s.Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY s.ProjectID, sd.sbjnum),
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
    FROM   SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum and s.Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY s.ProjectID, sd.sbjnum),
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
    FROM   SurveyData sd inner join survey s on sd.sbjnum = s.sbjnum and s.Created between '{req.StartDate}' and '{req.EndDate}' GROUP BY s.ProjectID, sd.sbjnum),
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
            string  Where = $"where s.ProjectID in ({req.ProjectId})";
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
		sd7.fieldId as FieldId7, sd7.fieldValue as FieldValue7,
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.created desc) as RowNum
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
 FieldValue7 as Status, convert(varchar, Created,101) asDate
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	inner join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)
	inner join ProjectFieldSample fs6 on cte.FieldId6 = fs6.FieldID and fs6.Code IN (cte.FieldValue6)
	inner join ProjectFieldSample fs7 on cte.FieldId7 = fs7.FieldID and fs7.Code IN (cte.FieldValue7)
    where RowNum = 1 and created between '{req.StartDate}' and '{req.EndDate}' select * from #Graph";


            if(string.IsNullOrEmpty(req.DistrictName))
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
	select  s.sbjnum, s.Created, s.ProjectID,
       
		sd2.fieldId as FieldId2, sd2.fieldValue as FieldValue2,
		sd3.fieldId as FieldId3, sd3.fieldValue as FieldValue3,
	   
		sd5.fieldId as FieldId5, sd5.fieldValue as FieldValue5,
	
	row_number() over (partition by  sd2.fieldId, sd2.fieldValue,sd3.fieldId,sd3.fieldValue ,sd5.fieldId,sd5.fieldValue order by s.created desc) as RowNum
	from survey s
		inner join SurveyData sd2 on s.sbjnum = sd2.sbjnum and sd2.FieldId in (50435, 50484, 55587) --District
		inner join SurveyData sd3 on s.sbjnum = sd3.sbjnum and sd3.FieldId in (50446, 50486, 55588)--Center close Survey Ids
		Inner join SurveyData sd5 on s.sbjnum = sd5.sbjnum and sd5.FieldId in (55592,50496,50635) {Where})
         	

select  convert(varchar, Created,101) asDate,  fs2.Title as District ,fs3.Title as Center, 
 FieldValue5 
 as Remarks
    into #Graph from cte
	inner join ProjectFieldSample fs2 on cte.FieldId2 = fs2.FieldID and fs2.Code IN (cte.FieldValue2)
	inner join ProjectFieldSample fs3 on cte.FieldId3 = fs3.FieldID and fs3.Code IN (cte.FieldValue3)
	Left join ProjectFieldSample fs5 on cte.FieldId5 = fs5.FieldID and fs5.Code IN (cte.FieldValue5)

    where RowNum = 1 and len(FieldValue5)  between 1 and 19 and created between '{req.StartDate}' and '{req.EndDate}' select * from #Graph

 
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
    }
}