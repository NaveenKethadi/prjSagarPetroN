using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace prjSagarPetroN.Models
{
    public class ItemShowModel
    {
        public string Item { get; set; }
        public string Units { get; set; }
        public string Description { get; set; }
        public decimal AvailableStockQty { get; set; }
        public decimal SafetyLevelQty { get; set; }
        public decimal Difference { get; set; }
        public decimal QtytobeProduced { get; set; }
        public string ProductionPlanningStatus { get; set; }
        public string ProductionPlanningDocNo { get; set; }
        public string TankMaster { get; set; }
        public decimal TankCapacity { get; set; }
        public decimal ClosingStock { get; set; }
        public decimal TopUpTankCapacity  { get; set; }
        public string Remarks { get; set; }
        public string DueDate { get; set; }
        public string ItemCode { get; set; }
        public int Itemid { get; set; }
        public int UnitId { get; set; }
        public string Warehouse { get; set; }
        public int Whid { get; set; }
        public decimal AvailableStockQty1 { get; set; }
        public decimal SafetyLevelQty1 { get; set; }
        public decimal Difference1 { get; set; }
        public decimal QtytobeProduced1 { get; set; }
        public decimal TopUpTankCapacity1 { get; set; }
        public string sVoucherNo { get; set; }

    }
    public class PostingData
    {
        public PostingData()
        {
            data = new List<Hashtable>();
        }
        public List<Hashtable> data { get; set; }
    }
}