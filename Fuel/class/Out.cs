using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fuel
{
    class Out
    {
        //номер карты    
        public string Card { get; set; }
        // точка обслуживания (АЗС)
        public string Azs { get; set; }
        // адрес АЗС
        public string AdressAzs { get; set; }
        // дата заправки
        public string DateFill { get; set; }
        // операция
        public string Operation { get; set; }
        // вид топлива
        public string TypeFuel { get; set; }
        // колличество 
        public string CountFuel { get; set; }
        // компания 
        public string NameCompany { get; set; }

        public Out(string c,string s, string a, string d, string o, string t, string co, string n)
        {
            Card = c;
            Azs = s;
            AdressAzs = a;
            DateFill = d;
            Operation = o;
            TypeFuel = t;
            CountFuel = co;
            NameCompany = n;
        }        

    }
}
