using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iCAR10p
{
    class Cliente
    {
        public string _rut { get; set; }
        public string _razon_social { get; set; }
        public string _nombre { get; set; }

        public Cliente (string rut, string razon_social)
        {
            _rut = rut;
            _razon_social = razon_social;
        }
        public Cliente(string rut, string razon_social, string nombre)
        {
            _rut = rut;
            _razon_social = razon_social;
            _nombre = nombre;
        }
    }
}
