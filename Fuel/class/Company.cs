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
       
        [DisplayName("Полное название фирмы")]
        [DataMember(Name = "Name")]
        public string Name { get; set; }
       
        [DisplayName("Полное наименование")]
        [DataMember(Name = "FullName")]
        public string FullName { get; set; }
        
        [DisplayName("как в Башнефть")]
        [DataMember(Name = "NameBash")]
        public string NameBash { get; set; }
        
        [DisplayName("как в Лукойл")]
        [DataMember(Name = "NameLuk")]
        public string NameLuk { get; set; }

        public Company(string n, string fn, string nb, string nl)
        {
            Name = n;
            FullName = fn;
            NameBash = nb;
            NameLuk = nl;
        }

    }
}
