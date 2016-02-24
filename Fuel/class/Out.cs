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
        public string ServicePoint { get; set; }
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

        public Out(string c,string s, string a, string d, string o, string t, string co)
        {
            Card = c;
            ServicePoint = s;
            AdressAzs = a;
            DateFill = d;
            Operation = o;
            TypeFuel = t;
            CountFuel = co;
        }        

    }
}
