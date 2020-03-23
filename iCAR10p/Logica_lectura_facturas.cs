using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;

namespace iCAR10p
{
    class Logica_lectura_facturas
    {
        public List<Factura> _facturas_leidas { get; set; }
        List<string> _patentes { get; set; }
        public TablaExcel _output_patentes { get; set; }
        

        private string _directorio_de_facturas;

        public Logica_lectura_facturas(string directorio)
        {
            _directorio_de_facturas = directorio;
            _facturas_leidas = new List<Factura>();
            _patentes = new List<string>();
        }

        public void Procesar_facturas()
        {

            var files = Directory.GetFiles(_directorio_de_facturas, "*.xml")
                     .Select(f => new ListViewItem(f))
                     .ToArray();

            foreach (var f in files)
            {
                string direccion = f.Text;
                Factura factura = new Factura();
                _facturas_leidas.Add(factura);

                string folio;
                int folio_int = 0;
                double _monto_neto = 0;
                double _iva;
                double _monto_total;
                string _rut_receptor;
                string _razon_social_receptor;
                string _fecha_emision;
                int _numero_patentes = 0;
                double _monto_por_patente;

                XmlTextReader xtr = new XmlTextReader(direccion);

                while (xtr.Read())
                {
                    if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "Folio")
                    {
                        folio = xtr.ReadElementContentAsString();
                        bool success = int.TryParse(folio, out folio_int);
                        factura._folio = folio_int;
                    }

                    else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "FchEmis")
                    {
                        _fecha_emision = xtr.ReadElementContentAsString();
                        factura._fecha_emision = _fecha_emision;
                    }

                    else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "RUTRecep")
                    {
                        _rut_receptor = xtr.ReadElementContentAsString();
                        factura._rut_receptor = _rut_receptor;
                    }

                    else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "RznSocRecep")
                    {
                        _razon_social_receptor = xtr.ReadElementContentAsString();
                        factura._razon_social_receptor = _razon_social_receptor;
                    }
                    else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "MntNeto")
                    {
                        _monto_neto = xtr.ReadElementContentAsDouble();
                        factura._monto_neto = _monto_neto;
                    }

                    else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "IVA")
                    {
                        _iva = xtr.ReadElementContentAsDouble();
                        factura._iva = _iva;
                    }

                    else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "MntTotal")
                    {
                        _monto_total = xtr.ReadElementContentAsDouble();
                        factura._monto_total = _monto_total;
                    }

                    if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DscItem")
                    {
                        string lector = xtr.ReadElementContentAsString();
                        string[] posibles_patentes = lector.Split(',');

                        for (int i = 0; i < posibles_patentes.Length; i++)
                        {
                            string posible_patente = Regex.Replace(posibles_patentes[i], @"[^a-zA-Z0-9\-]", "");

                            if (posible_patente.Length == 8)
                            {
                                posible_patente = posible_patente.Substring(0, 6);
                                _patentes.Add(posible_patente);
                                factura._patentes_factura.Add(posible_patente);
                                _numero_patentes++;
                            }
                            else if (posible_patente.Length == 6)
                            {
                                _patentes.Add(posible_patente);
                                factura._patentes_factura.Add(posible_patente);
                                _numero_patentes++;
                            }
                            else
                            {
                                string error = xtr.BaseURI;
                            }
                        }

                    }

                    
                }

                if (_numero_patentes > 0)
                {
                    _monto_por_patente = _monto_neto / _numero_patentes;
                    _monto_por_patente = Math.Round(_monto_por_patente, 1);
                    factura._monto_por_patente = _monto_por_patente;
                }
            }

            
        }

        public void Generar_Tablas_Excel()
        {
            _output_patentes = new TablaExcel("Patentes_facturadas");
            _output_patentes._directorio = _directorio_de_facturas;

            Dictionary<int, Dictionary<string, string>> _output_patentes_datos = new Dictionary<int, Dictionary<string, string>>();

            //Se completa el diccionario con los nombres de las columnas de la TablaExcel output
            _output_patentes._nombre_columnas_tabla.Add(0, "Patente");
            _output_patentes._nombre_columnas_tabla.Add(1, "Monto_facturado");
            _output_patentes._nombre_columnas_tabla.Add(2, "Folio_factura");
            _output_patentes._nombre_columnas_tabla.Add(3, "Cliente");
            _output_patentes._nombre_columnas_tabla.Add(4, "Rut_Cliente");

            //Comenzamos con la primera fila de las titulos
            Dictionary<string, string> columnas = new Dictionary<string, string>();
            columnas.Add("Patente", "Patente");
            columnas.Add("Monto_facturado", "Monto_facturado");
            columnas.Add("Folio_factura", "Folio_factura");
            columnas.Add("Cliente", "Cliente");
            columnas.Add("Rut_Cliente", "Rut_Cliente");

            //Se agrega primera fila
            int contador_filas = 0;
            _output_patentes_datos.Add(contador_filas, columnas);
            contador_filas++;

            foreach(Factura mi_factura in _facturas_leidas)
            {
                for (int i = 0; i < mi_factura._patentes_factura.Count; i++)
                {
                    Dictionary<string, string> fila_i = new Dictionary<string, string>();

                    //Patente
                    fila_i.Add("Patente", _patentes[i]);
                    // Monto_facturado
                    fila_i.Add("Monto_facturado", mi_factura._monto_por_patente.ToString());

                    //Folio_factura
                    fila_i.Add("Folio_factura", mi_factura._folio.ToString());

                    //Cliente
                    fila_i.Add("Cliente", mi_factura._razon_social_receptor);

                    //Rut_Cliente
                    fila_i.Add("Rut_Cliente", mi_factura._rut_receptor);

                    _output_patentes_datos.Add(contador_filas, fila_i);
                    contador_filas++;
                }
            }

            

            _output_patentes._datos_tabla = _output_patentes_datos;
        }
    }
}
