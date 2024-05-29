using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace prjSagarPetroN.Models
{
    public class BomModel
    {
        public int iBOM { get; set; }
        public int iVersion { get; set; }
        public string BomName { get; set; }
        public int ItemId { get; set; }
        public string Item { get; set; }
        public int iProductType { get; set; }
        public decimal fQty { get; set; }
        public int bInput { get; set; }
    }
}