using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChickenFarmExcel
{
    public class FinalData
    {
        public string No { get; set; }
        public string CarNo { get; set; }
        public DateTime? DepartureTime { get; set; }
        public DateTime? ReturnTime { get; set; }
        public string ChickenFarm { get; set; } 
        public string Slaughterhouse { get; set; }
        public DateTime? Come1 { get; set; }
        public DateTime? Out1 { get; set; }
        public DateTime? Come2 { get; set; }
        public DateTime? Out2 { get; set; }
    }
}
