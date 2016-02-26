using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fuel
{
    //[JsonObject(MemberSerialization.OptIn)]
    class cellExcel
    {

        // [JsonProperty("CellCard")]
        public int CellCard { get; set; }
        public int CellAzs{ get; set; }        
        public int CellCompany { get; set; }
        public int CellAdressAzs { get; set; }
        public int CellDateFill { get; set; }
        public int CellOperation { get; set; }
        public int CellFuelT { get; set; }
        public int CellCountF { get; set; }
        public int FirstRow { get; set; }
        public int LastRow { get; set; }
        public string FolderPatch { get; set; }
        public int FolderMonth { get; set; }
        public int ListExl { get; set; }

    }
}
