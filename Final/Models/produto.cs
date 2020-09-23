using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Final.Models
{
    public class produto
    {
        public int Id { get; set; }
        public int CodProud { get; set; }

        public string DescProduto { get; set; }

        public string IVA { get; set; }

        public string ICMS1 { get; set; }

        public string ICMS2 { get; set; }

        public int CodICM { get; set; }

        public string PercBaSerDST { get; set; }

        public string PercBaSerd { get; set; }
        public string tipo { get; set; }
        public string NCM { get; set; }

        public string RegTribEstadual { get; set; }
        public string BaseICMS { get; set; }

        public string VendaIntCred { get; set; }

        public string VendaIntDeb { get; set; }

        public string ICMCred { get; set; }

        public string ICMDeb { get; set; }

        public string MVAind { get; set; }

        public string MVAatac { get; set; }

        public string MVA4 { get; set; }

        public string MVA7 { get; set; }

        public string MVA12 { get; set; }
    }
}