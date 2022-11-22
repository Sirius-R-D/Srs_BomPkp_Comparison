using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Srs_BomPkp_Comparison
{
    class XlsxFile
    {
        public List<Components> component = new List<Components>();
    }
    public class Components
    {
        public string CustomerStockCode;
        public string SiriusStockCode;
        public string StockName;
        public string PartNo;
        public string Refference;
    }
}
