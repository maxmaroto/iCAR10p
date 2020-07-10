using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iCAR10p
{
    class Detalle_Factura
    {
        public int _numero_linea_detalle { get; set; }
        public string _tipo_codigo { get; set; }
        public string _valor_codigo { get; set; }
        public string _nombre_item { get; set; }
        public string _descripcion_item { get; set; }
        public double _cantidad_item { get; set; }
        public string _tipo_unidades_item { get; set; }
        public double _precio_item { get; set; }
        public double _monto_item { get; set; }
        public List<string> _patentes_detalle { get; set; }

        public Detalle_Factura(int numero_detalle)
        {
            _numero_linea_detalle = numero_detalle;
            _patentes_detalle = new List<string>();
        }
    }
}
