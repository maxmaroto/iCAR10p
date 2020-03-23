﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iCAR10p
{
    class Factura
    {
        public int _folio { get; set; }
        public double _monto_neto { get; set; }
        public double _iva { get; set; }
        public double _monto_total { get; set; }
        public string _rut_receptor { get; set; }
        public string _razon_social_receptor { get; set; }
        public string _fecha_emision { get; set; }

        public int _numero_patentes { get; set; }
        public double _monto_por_patente { get; set; }

        public List<string> _patentes_factura { get; set; }

        public Factura()
        {
            _patentes_factura = new List<string>();
        }


    }
}
