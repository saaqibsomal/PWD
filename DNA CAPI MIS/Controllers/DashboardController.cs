using DNA_CAPI_MIS.Models;
using DNA_CAPI_MIS.Service;
using System;
using System.Collections.Generic;
using System.Linq;
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
                sdpStatus  = service.SDPStatus(req);
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
                sdpStatus  = service.StuffDetailReportData(req);
            }
            catch (Exception ex)
            {
                throw;
            }
            return Json(sdpStatus, JsonRequestBehavior.AllowGet);
        }
    }
}