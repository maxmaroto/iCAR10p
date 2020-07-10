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
        public TablaExcel _output_facturas { get; set; }


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
                int _contador_detalle = 0;

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

                    if(xtr.NodeType == XmlNodeType.Element && xtr.Name == "Detalle")
                    {
                        _contador_detalle++;
                        Detalle_Factura detalle = new Detalle_Factura(_contador_detalle);
                        detalle._tipo_unidades_item = ""; //PARCHE ORDINARIO HARDCODE

                        factura._lineas_de_detalle.Add(detalle);

                        xtr.Read();

                        string asd = xtr.Value.ToString();

                        int contador_intentos = 0;

                        while(xtr.Name != "Detalle" && contador_intentos < 1000)
                        {
                            contador_intentos++;

                            if(contador_intentos == 1000)
                            {
                                int folilio = factura._folio;
                                int adkuh = 0;
                            }

                            if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "NroLinDet")
                            {
                                int NroLinDet = Int32.Parse(xtr.ReadElementContentAsString());

                                if (detalle._numero_linea_detalle != NroLinDet)
                                {
                                    int a = 0;
                                }
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "TpoCodigo")
                            {
                                detalle._tipo_codigo = xtr.ReadElementContentAsString();
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "VlrCodigo")
                            {
                                detalle._valor_codigo = xtr.ReadElementContentAsString();
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "NmbItem")
                            {
                                detalle._nombre_item = xtr.ReadElementContentAsString();
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "DscItem")
                            {
                                detalle._descripcion_item = xtr.ReadElementContentAsString();

                                string lector = detalle._descripcion_item;
                                char[] separador = { ',', '-', '/', '\n', '\r' };
                                string[] posibles_patentes = lector.Split(separador);

                                for (int i = 0; i < posibles_patentes.Length; i++)
                                {
                                    string posible_patente = Regex.Replace(posibles_patentes[i], @"[^a-zA-Z0-9\-]", "");

                                    if (posible_patente.Length == 8)
                                    {
                                        posible_patente = posible_patente.Substring(0, 6);
                                        //_patentes.Add(posible_patente);
                                        detalle._patentes_detalle.Add(posible_patente);
                                        _numero_patentes++;
                                    }
                                    else if (posible_patente.Length == 6)
                                    {
                                        //_patentes.Add(posible_patente);
                                        detalle._patentes_detalle.Add(posible_patente);
                                        _numero_patentes++;
                                    }
                                    else
                                    {
                                        string error = xtr.BaseURI;
                                    }
                                }
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "QtyItem")
                            {
                                detalle._cantidad_item = xtr.ReadElementContentAsDouble();
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "UnmdItem")
                            {
                                detalle._tipo_unidades_item = xtr.ReadElementContentAsString();
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "PrcItem")
                            {
                                detalle._precio_item = xtr.ReadElementContentAsDouble();
                            }

                            else if (xtr.NodeType == XmlNodeType.Element && xtr.Name == "MontoItem")
                            {
                                detalle._monto_item = xtr.ReadElementContentAsDouble();
                            }

                           
                            xtr.Read();
                            asd = xtr.Value.ToString();
                            
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

        public void Generar_Tablas_Excel_Patentes()
        {
            _output_patentes = new TablaExcel("Patentes_facturadas");
            _output_patentes._directorio = _directorio_de_facturas;

            Dictionary<int, Dictionary<string, string>> _output_patentes_datos = new Dictionary<int, Dictionary<string, string>>();

            //Se completa el diccionario con los nombres de las columnas de la TablaExcel output
            _output_patentes._nombre_columnas_tabla.Add(0, "Patente");
            _output_patentes._nombre_columnas_tabla.Add(1, "Monto_Item");
            _output_patentes._nombre_columnas_tabla.Add(2, "Folio_factura");
            _output_patentes._nombre_columnas_tabla.Add(3, "Cliente");
            _output_patentes._nombre_columnas_tabla.Add(4, "Rut_Cliente");
            _output_patentes._nombre_columnas_tabla.Add(5, "N_Detalle");
            _output_patentes._nombre_columnas_tabla.Add(6, "Tipo_Codigo");
            _output_patentes._nombre_columnas_tabla.Add(7, "Valor_Codigo");
            _output_patentes._nombre_columnas_tabla.Add(8, "Nombre_Item");
            _output_patentes._nombre_columnas_tabla.Add(9, "Descripcion_Item");
            _output_patentes._nombre_columnas_tabla.Add(10, "Cantidad");
            _output_patentes._nombre_columnas_tabla.Add(11, "Tipo_Unidades");

            //Comenzamos con la primera fila de las titulos
            Dictionary<string, string> columnas = new Dictionary<string, string>();
            columnas.Add("Patente", "Patente");
            columnas.Add("Monto_Item", "Monto_Item");
            columnas.Add("Folio_factura", "Folio_factura");
            columnas.Add("Cliente", "Cliente");
            columnas.Add("Rut_Cliente", "Rut_Cliente");
            columnas.Add("N_Detalle", "N_Detalle");
            columnas.Add("Tipo_Codigo", "Tipo_Codigo");
            columnas.Add("Valor_Codigo", "Valor_Codigo");
            columnas.Add("Nombre_Item", "Nombre_Item");
            columnas.Add("Descripcion_Item", "Descripcion_Item");
            columnas.Add("Cantidad", "Cantidad");
            columnas.Add("Tipo_Unidades", "Tipo_Unidades");
            //Se agrega primera fila
            int contador_filas = 0;
            _output_patentes_datos.Add(contador_filas, columnas);
            contador_filas++;

            foreach(Factura mi_factura in _facturas_leidas)
            {
                foreach (Detalle_Factura mi_detalle in mi_factura._lineas_de_detalle)
                {
                    for (int i = 0; i < mi_detalle._patentes_detalle.Count; i++)
                    {
                        Dictionary<string, string> fila_i = new Dictionary<string, string>();

                        //Patente
                        fila_i.Add("Patente", mi_detalle._patentes_detalle[i]);
                        // Monto_facturado
                        fila_i.Add("Monto_Item", mi_detalle._precio_item.ToString());
                        //Folio_factura
                        fila_i.Add("Folio_factura", mi_factura._folio.ToString());
                        //Cliente
                        fila_i.Add("Cliente", mi_factura._razon_social_receptor);
                        //Rut_Cliente
                        fila_i.Add("Rut_Cliente", mi_factura._rut_receptor);


                        //N_Detalle
                        fila_i.Add("N_Detalle", mi_detalle._numero_linea_detalle.ToString());
                        // Tipo_Codigo
                        fila_i.Add("Tipo_Codigo", mi_detalle._tipo_codigo);
                        //Valor_Codigo
                        fila_i.Add("Valor_Codigo", mi_detalle._valor_codigo);
                        //Nombre_Item
                        fila_i.Add("Nombre_Item", mi_detalle._nombre_item);
                        //Descripcion_Item
                        fila_i.Add("Descripcion_Item", mi_detalle._descripcion_item);
                        //Cantidad
                        fila_i.Add("Cantidad", mi_detalle._cantidad_item.ToString());
                        //Tipo_Unidades
                        fila_i.Add("Tipo_Unidades", mi_detalle._tipo_unidades_item);
                        
                        _output_patentes_datos.Add(contador_filas, fila_i);
                        contador_filas++;
                    }
                }
            }

            

            _output_patentes._datos_tabla = _output_patentes_datos;
        }

        public void Generar_Tablas_Excel_Facturas()
        {
            _output_facturas = new TablaExcel("Facturas_procesadas");
            _output_facturas._directorio = _directorio_de_facturas;

            Dictionary<int, Dictionary<string, string>> _output_facturas_datos = new Dictionary<int, Dictionary<string, string>>();

            //Se completa el diccionario con los nombres de las columnas de la TablaExcel output
            _output_facturas._nombre_columnas_tabla.Add(0, "Folio_factura");
            _output_facturas._nombre_columnas_tabla.Add(1, "Fecha");
            _output_facturas._nombre_columnas_tabla.Add(2, "Cliente");
            _output_facturas._nombre_columnas_tabla.Add(3, "Rut_Cliente");
            _output_facturas._nombre_columnas_tabla.Add(4, "Monto_neto");

            /*
            _output_patentes._nombre_columnas_tabla.Add(3, "Cliente");
            _output_patentes._nombre_columnas_tabla.Add(4, "Rut_Cliente");
            _output_patentes._nombre_columnas_tabla.Add(5, "N_Detalle");
            _output_patentes._nombre_columnas_tabla.Add(6, "Tipo_Codigo");
            _output_patentes._nombre_columnas_tabla.Add(7, "Valor_Codigo");
            _output_patentes._nombre_columnas_tabla.Add(8, "Nombre_Item");
            _output_patentes._nombre_columnas_tabla.Add(9, "Descripcion_Item");
            _output_patentes._nombre_columnas_tabla.Add(10, "Cantidad");
            _output_patentes._nombre_columnas_tabla.Add(11, "Tipo_Unidades");*/

            //Comenzamos con la primera fila de las titulos
            Dictionary<string, string> columnas = new Dictionary<string, string>();
            columnas.Add("Folio_factura", "Folio_factura");
            columnas.Add("Fecha", "Fecha");
            columnas.Add("Cliente", "Cliente");
            columnas.Add("Rut_Cliente", "Rut_Cliente");
            columnas.Add("Monto_neto", "Monto_neto");

            /*
            columnas.Add("Rut_Cliente", "Rut_Cliente");
            columnas.Add("N_Detalle", "N_Detalle");
            columnas.Add("Tipo_Codigo", "Tipo_Codigo");
            columnas.Add("Valor_Codigo", "Valor_Codigo");
            columnas.Add("Nombre_Item", "Nombre_Item");
            columnas.Add("Descripcion_Item", "Descripcion_Item");
            columnas.Add("Cantidad", "Cantidad");
            columnas.Add("Tipo_Unidades", "Tipo_Unidades");
            */
            //Se agrega primera fila
            int contador_filas = 0;
            _output_facturas_datos.Add(contador_filas, columnas);
            contador_filas++;

            foreach (Factura mi_factura in _facturas_leidas)
            {

                Dictionary<string, string> fila_i = new Dictionary<string, string>();


                //Folio_factura
                fila_i.Add("Folio_factura", mi_factura._folio.ToString());
                //Fecha
                fila_i.Add("Fecha", mi_factura._fecha_emision.ToString());
                //Cliente
                fila_i.Add("Cliente", mi_factura._razon_social_receptor);
                //Rut_Cliente
                fila_i.Add("Rut_Cliente", mi_factura._rut_receptor);
                //Monto_neto
                fila_i.Add("Monto_neto", mi_factura._monto_neto.ToString());

                _output_facturas_datos.Add(contador_filas, fila_i);
                contador_filas++;

            }



            _output_facturas._datos_tabla = _output_facturas_datos;
        }
    }
}
