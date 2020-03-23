using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iCAR10p
{
    class TablaExcel
    {
        public Dictionary<int, String> _nombre_columnas_tabla { get; set; }
        public string _nombre_tabla { get; set; }
        public Dictionary<int, Dictionary<string, string>> _datos_tabla { get; set; }
        public string _directorio { get; set; }

        public TablaExcel (string nombre)
        {
            _nombre_tabla = nombre;
            _nombre_columnas_tabla = new Dictionary<int, string>();
            _datos_tabla = new Dictionary<int, Dictionary<string, string>>();
        }


    }
}
