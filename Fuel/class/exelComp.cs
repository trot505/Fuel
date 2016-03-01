using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fuel
{
    //[JsonObject(MemberSerialization.OptIn)]
    class exelComp
    {
        // [JsonProperty("CellCard")]
        public int BriefName { get; set; }
        public int FullName { get; set; }
        public int BashName { get; set; }
        public int LukName { get; set; }
        public int RangeFirst { get; set; }
        public int RangeLast { get; set; }
        public int ListPage { get; set; }
    }
}
