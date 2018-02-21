using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Agenty_WF
{
    class YRdb
    {
        //public int Id;
        //public string Agent;
        public string Vlice;
        public string Osnovanie;
        public string Dogovor;
        //public string DataDog;
        public string Podpisant;

        public YRdb(string Vlice, string Osnovanie, string Dogovor, string Podpisant)
        {
            //this.Id = Id;
            this.Vlice = Vlice;
            this.Osnovanie = Osnovanie;
            this.Dogovor = Dogovor;
            //this.DataDog = DataDog;
            this.Podpisant = Podpisant;
        }

    }
}
