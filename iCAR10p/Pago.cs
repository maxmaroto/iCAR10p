using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iCAR10p
{
    class Pago
    {
        public bool _pago_es_contabilizable { get; set; }
        public int _id_pago { get; set; }
        public int _monto_pago { get; set; }
        public string _fecha_pago { get; set; }

        public Pago(int id, int _monto_primer_pago)
        {
            _id_pago = id;
            _monto_pago = _monto_primer_pago;
            _pago_es_contabilizable = true;
        }
    }
}
