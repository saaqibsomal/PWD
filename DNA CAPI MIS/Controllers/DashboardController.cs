using DNA_CAPI_MIS.Models;
using DNA_CAPI_MIS.Service;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;

namespace DNA_CAPI_MIS.Controllers
{
    public class DashboardController : Controller
    {
        DashboardService service = new DashboardService();

        [HttpPost]
        public JsonResult GetSdpsStatus(DashboardRequest req)
        {
            List<SDPsStatus> sdpStatus = new List<SDPsStatus>();
            try
            {
                sdpStatus = service.SDPStatus(req);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(sdpStatus, JsonRequestBehavior.AllowGet);
        }



        [HttpPost]
        public JsonResult StaffPosition(DashboardRequest req)
        {
            List<StuffPosition> sdpStatus = new List<StuffPosition>();
            try
            {
                sdpStatus = service.StuffDetailReportData(req);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(sdpStatus, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult ContraceptiveStockPosition(DashboardRequest req)
        {
            ContraceptiveStockPositionModel sdpStatus = new ContraceptiveStockPositionModel();
            try
            {
                sdpStatus = service.ContraceptiveStockPosition(req);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(sdpStatus, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult MonitoringVisitsReport(DashboardRequest req)
        {
            List<PdfDetailReport> report = new List<PdfDetailReport>();
            try
            {
                var name = string.Empty;
                IPrincipal users = null;
                try
                {
                    name = User?.Identity?.Name;
                    users = User;
                }
                catch (Exception ex)
                {

                }
                report = service.MonitoringVisitsReport(req, name, users);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(report, JsonRequestBehavior.AllowGet);
        }


        [HttpPost]
        public JsonResult StockOfContraceptive(DashboardRequest req)
        {
            List<StockOfContraceptiveResponse> report = new List<StockOfContraceptiveResponse>();
            try
            {

                report = service.StockOfConteraceptives(req);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(report, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult DetailStockOfContraceptive(DashboardRequest req)
        {
            List<ContraceptiveStock> report = new List<ContraceptiveStock>();
            try
            {
                report = service.DetailOfContraceptive(req);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(report, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public JsonResult Report(string sbjnum, string Heading)
        {
            string report = string.Empty;
            try
            {
                string baseUrl = HttpContext?.Request?.Url?.GetLeftPart(UriPartial.Authority);
                report = service.GetReport(sbjnum, Heading, baseUrl);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(report, JsonRequestBehavior.AllowGet);
        }


        [HttpGet]
        public JsonResult OfficerVisitReport()
        {
            List<SurveyReportViewModel> report = new List<SurveyReportViewModel>();
            try
            {

                report = service.OfficerVisitReport();
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(report, JsonRequestBehavior.AllowGet);
        }
    }
}