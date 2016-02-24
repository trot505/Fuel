using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Fuel {
    [DataContract]
    class Company
    {
        [Category("Name")]
        [DisplayName("Полное название фирмы")]
        [DataMember(Name = "Name")]
        public string Name { get; set; }
        [Category("Name")]
        [DisplayName("как в Башнефть")]
        [DataMember(Name = "NameBash")]
        public string NameBash { get; set; }
        [Category("Name")]
        [DisplayName("как в Лукойл")]
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
