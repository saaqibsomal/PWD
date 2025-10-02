using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DNA_CAPI_MIS.Models
{
    public class ContraceptiveStockPositionResponse
    {
        public string Date { get; set; }
        public string SDP { get; set; }
        public string District { get; set; }
        public string Center { get; set; }
        public int CondomsStock { get; set; }
        public int POP { get; set; }
        public int COC { get; set; }
        public int ECP { get; set; }
        public int ThreemonthsInj { get; set; }
        public int DefoStock { get; set; }
        public int IUD { get; set; }
        public int Jadelle { get; set; }

    }

    public class ContraceptiveStockPositionModel
    {
        public List<ContraceptiveStockPositionResponse> contraceptiveStockPositionResponse { get; set; } = new List<ContraceptiveStockPositionResponse>();
        public List<ContraceptiveStockPositionStockWise> contraceptiveStockPositionStockWise { get; set; } = new List<ContraceptiveStockPositionStockWise>(); 
        public List<CenterStockAlert>  centerStockAlerts { get; set; } = new List<CenterStockAlert>(); 
    }

    public class ContraceptiveStockPositionStockWise
    {
 
        public string SDP { get; set; }
        public int CondomsStock { get; set; }
        public int POP { get; set; }
        public int COC { get; set; }
        public int ECP { get; set; }
        public int ThreemonthsInj { get; set; }
        public int DefoStock { get; set; }
        public int IUD { get; set; }
        public int Jadelle { get; set; }

    }

    public class CenterStockAlert
    {
        public string Center { get; set; }
        public string District { get; set; }
        public string Message { get; set; }
        public string AlertLevel { get; set; }
        public int TotalStock { get; set; }    // Added for sorting
    }
}