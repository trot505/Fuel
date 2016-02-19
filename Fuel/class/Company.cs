using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Fuel {
    [DataContract]
    class Company
    {
        [DataMember(Name = "Name")]
        public string Name { get; set; }
        [DataMember(Name = "NameBash")]
        public string NameBash { get; set; }
        [DataMember(Name = "NameLuk")]
        public string NameLuk { get; set; }

        public Company(string n, string nb, string nl)
        {
            Name = n;
            NameBash = nb;
            NameLuk = nl;
        }

    }
}
