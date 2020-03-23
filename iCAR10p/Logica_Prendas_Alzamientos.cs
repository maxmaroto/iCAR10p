using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace iCAR10p
{
    class Logica_Prendas_Alzamientos
    {
        public string _nombre_archivo_a_contabilizar { get; set; }

        public string _directorio {get; set;}

        public TablaExcel output_ExcelTabla_a_contabilizar { get; set; }
        public TablaExcel output_ExcelTabla_con_problemas_para_contabilizar { get; set; }

        public Dictionary<int, Dictionary<string, string>> _output_datos_a_contabilizar { get; set; }
        public Dictionary<int, Dictionary<string, string>> _output_datos_con_problemas_para_contabilizar { get; set; }

        private Dictionary<string, int> _pagos_y_montos { get; set; }

        private TablaExcel _input_datos;

        private List<Cliente> _mis_clientes;

        private List<Pago> _mis_pagos;

        private bool _todos_los_datos_pueden_ser_contabilidados { get; set; }

        public Logica_Prendas_Alzamientos(TablaExcel mis_datos)
        {
            output_ExcelTabla_a_contabilizar = new TablaExcel("Output");            
            output_ExcelTabla_con_problemas_para_contabilizar = new TablaExcel("Output operaciones no contabilizables");
            
            _mis_clientes = new List<Cliente>();
            _mis_pagos = new List<Pago>();
            _input_datos = mis_datos;
            _todos_los_datos_pueden_ser_contabilidados = true;

            Cliente mundo_credito = new Cliente("76224981-2", "SERVICIOS FINANCIEROS MUNDO CREDITO S.A.", "MUNDO CREDITO");
            Cliente GMF = new Cliente("94050000-1", "GENERAL MOTORS FINANCIAL CHILE S.A.","GMF");

            _mis_clientes.Add(mundo_credito);
            _mis_clientes.Add(GMF);          

            

        }

        public void Procesar_datos()
        {
            _output_datos_a_contabilizar = new Dictionary<int, Dictionary<string, string>>();
            _output_datos_con_problemas_para_contabilizar = new Dictionary<int, Dictionary<string, string>>();
            int contador_filas_contabilizadas = 0;
            int contador_filas_no_contabilizadas = 0;
            output_ExcelTabla_a_contabilizar._directorio = _directorio;
            output_ExcelTabla_con_problemas_para_contabilizar._directorio = _directorio;

            #region Comenzamos con la primera fila de las titulos
            Dictionary<string, string> columnas = new Dictionary<string, string>();
            columnas.Add("Linea", "Linea");
            columnas.Add("Cuenta", "Cuenta");
            columnas.Add("Desc.Cuenta", "Desc.Cuenta");
            columnas.Add("Debe", "Debe");
            columnas.Add("Haber", "Haber");
            columnas.Add("Glosa", "Glosa");
            columnas.Add("Tipo Entidad", "Tipo Entidad");
            columnas.Add("Entidad", "Entidad");
            columnas.Add("Nombre Entidad", "Nombre Entidad");
            columnas.Add("Tipo Docto", "Tipo Docto");
            columnas.Add("Emisor Docto", "Emisor Docto");
            columnas.Add("Nro Docto", "Nro Docto");
            columnas.Add("Correlativo Docto", "Correlativo Docto");
            columnas.Add("Fecha Vcto", "Fecha Vcto");
            columnas.Add("Tipo Linea", "Tipo Linea");
            columnas.Add("Centro de Costo", "Centro de Costo");
            columnas.Add("Sucursal", "Sucursal");
            columnas.Add("N° Doc.Pagado", "N° Doc.Pagado");
            columnas.Add("Folio Fact.", "Folio Fact.");
            columnas.Add("Rut Comprador", "Rut Comprador");
            columnas.Add("Rut Municipalidad", "Rut Municipalidad");
            columnas.Add("Id.I-Car", "Id.I-Car");
            columnas.Add("Tipo Operación", "Tipo Operación");
            columnas.Add("Placa Patente", "Placa Patente");

            _output_datos_a_contabilizar.Add(contador_filas_contabilizadas, columnas);
            contador_filas_contabilizadas++;
            _output_datos_con_problemas_para_contabilizar.Add(contador_filas_no_contabilizadas, columnas);
            contador_filas_no_contabilizadas++;

            //Se completa el diccionario con los nombres de las columnas de la TablaExcel output
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(0, "Linea");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(1, "Cuenta");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(2, "Desc.Cuenta");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(3, "Debe");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(4, "Haber");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(5, "Glosa");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(6, "Tipo Entidad");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(7, "Entidad");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(8, "Nombre Entidad");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(9, "Tipo Docto");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(10, "Emisor Docto");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(11, "Nro Docto");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(12, "Correlativo Docto");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(13, "Fecha Vcto");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(14, "Tipo Linea");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(15, "Centro de Costo");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(16, "Sucursal");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(17, "N° Doc.Pagado");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(18, "Folio Fact.");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(19, "Rut Comprador");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(20, "Rut Municipalidad");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(21, "Id.I-Car");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(22, "Tipo Operación");
            output_ExcelTabla_a_contabilizar._nombre_columnas_tabla.Add(23, "Placa Patente");


            #endregion

            for (int i = 1; i < _input_datos._datos_tabla.Count; i++)
            {
                Dictionary<string, string> fila_datos = _input_datos._datos_tabla[i];

                //Cambiamos el nombre de los archivos
                if (i == 1)
                {
                    output_ExcelTabla_a_contabilizar._nombre_tabla = convertidor_fecha(fila_datos["Fecha"])+" - " + "Prendas y alzamientos";
                    output_ExcelTabla_con_problemas_para_contabilizar._nombre_tabla = convertidor_fecha(fila_datos["Fecha"]) + " - " + "Prendas y alzamientos no contabilizados";
                }

                //Vemos si la fila contiene información de un nuevo pago, de ser así se crea el objeto pago
                int id_pago;
                bool id_pago_es_numero = int.TryParse(fila_datos["PAGO"], out id_pago);

                int monto_pago;
                bool monto_pago_es_numero = int.TryParse(fila_datos["TOTAL"], out monto_pago);

                bool pago_ya_existe = false;

                foreach (Pago _mi_pago in _mis_pagos)
                {
                    if (_mi_pago._id_pago == id_pago)
                    {
                        _mi_pago._monto_pago += monto_pago;
                        pago_ya_existe = true;
                    }
                }

                if (!pago_ya_existe)
                {
                    if (id_pago_es_numero && monto_pago_es_numero)
                    {
                        Pago nuevo_pago = new Pago(id_pago, monto_pago);
                        string fecha = "";
                        fecha = convertidor_fecha(fila_datos["Fecha"]);
                        nuevo_pago._fecha_pago = fecha;
                        _mis_pagos.Add(nuevo_pago);
                    }
                    else
                    {
                        int a = 0;
                    }
                }

                //Validamos que el pago tenga siempre una id de icar asociada, sino no se contabiliza
                string id_icar = fila_datos["ID SOLICITUD"];
                int id;
                bool isNumeric = int.TryParse(id_icar, out id);
                if (!isNumeric)
                {
                    foreach (Pago _mi_pago in _mis_pagos)
                    {
                        if (_mi_pago._id_pago == id_pago)
                        {
                            _mi_pago._pago_es_contabilizable = false;
                            _todos_los_datos_pueden_ser_contabilidados = false;
                        }
                    }
                }
            }

            if (!_todos_los_datos_pueden_ser_contabilidados)
            {
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(0, "Linea");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(1, "Cuenta");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(2, "Desc.Cuenta");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(3, "Debe");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(4, "Haber");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(5, "Glosa");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(6, "Tipo Entidad");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(7, "Entidad");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(8, "Nombre Entidad");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(9, "Tipo Docto");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(10, "Emisor Docto");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(11, "Nro Docto");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(12, "Correlativo Docto");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(13, "Fecha Vcto");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(14, "Tipo Linea");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(15, "Centro de Costo");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(16, "Sucursal");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(17, "N° Doc.Pagado");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(18, "Folio Fact.");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(19, "Rut Comprador");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(20, "Rut Municipalidad");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(21, "Id.I-Car");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(22, "Tipo Operación");
                output_ExcelTabla_con_problemas_para_contabilizar._nombre_columnas_tabla.Add(23, "Placa Patente");


            }

            for (int i = 1; i < _input_datos._datos_tabla.Count; i++)
            {
                Dictionary<string, string> fila_datos = _input_datos._datos_tabla[i];
                Dictionary<string, string> fila_i = new Dictionary<string, string>();
                bool contabilizar = true;

                int id_pago_fila;
                int id_pago = 0;
                bool id_pago_es_numero = int.TryParse(fila_datos["PAGO"], out id_pago_fila);
                foreach (Pago _mi_pago in _mis_pagos)
                {
                    if (_mi_pago._id_pago == id_pago_fila)
                    {
                        
                        if (!_mi_pago._pago_es_contabilizable)
                        {
                            contabilizar = false;
                        }
                        break;
                    }
                    id_pago++;

                }
                #region Llenado de datos
                //Linea
                fila_i.Add("Linea", "");
                // Cuenta
                string cuenta = "216124";
                fila_i.Add("Cuenta", cuenta);

                //Desc. cuenta
                fila_i.Add("Desc.Cuenta", "");

                // Debe
                fila_i.Add("Debe", fila_datos["TOTAL"]);

                //Haber
                fila_i.Add("Haber", "0");

                //Glosa se aprovecha de setear la fecha
                string fecha = "";
                
                fecha = convertidor_fecha(fila_datos["Fecha"]);
                string glosa = fecha + " Pago: " + fila_datos["PAGO"] + " por " + _mis_pagos[id_pago]._monto_pago.ToString() + " Pago TX" + fila_datos["PRENDA O ALZAMIENTO"] + " " + fila_datos["Patente"];
                fila_i.Add("Glosa", glosa);

                //Tipo Entidad
                fila_i.Add("Tipo Entidad", "CLIENTE");

                //Entidad y nombre entidad
                string entidad = "";
                string nombre_entidad = "";
                foreach(Cliente mi_cliente in _mis_clientes)
                {
                    if (fila_datos["ACREEDOR"] == mi_cliente._nombre)
                    {
                        entidad = "00" + mi_cliente._rut;
                        nombre_entidad = mi_cliente._razon_social;
                        break;
                    }
                }
                fila_i.Add("Entidad", entidad);
                fila_i.Add("Nombre Entidad", "");

                //Tipo Documento
                string tipo_documento = "TX"+ fila_datos["PRENDA O ALZAMIENTO"];
                fila_i.Add("Tipo Docto", tipo_documento);

                //Emisor Documento
                fila_i.Add("Emisor Docto", "");

                //Numero documento
                string numero_documento = fila_datos["N° REPERTORIO"];
                fila_i.Add("Nro Docto", numero_documento);

                //Correlativo documento
                fila_i.Add("Correlativo Docto", "");

                //Fecha
                fila_i.Add("Fecha Vcto", fecha);

                //Tipo linea
                fila_i.Add("Tipo Linea", "");

                //Centro de costo
                fila_i.Add("Centro de Costo", "");

                //Sucursal
                fila_i.Add("Sucursal", "");

                //N° Doc.Pagado
                fila_i.Add("N° Doc.Pagado", "");

                //Folio de factura, en este caso es el repertorio, ya que no hay factura
                fila_i.Add("Folio Fact.", fila_datos["N° REPERTORIO"]);

                //Rut comprador
                string rut_comprador = fila_datos["RUT"];
                fila_i.Add("Rut Comprador", rut_comprador);

                //Rut municipalidad
                fila_i.Add("Rut Municipalidad", "");

                //Id i-CAR
                string id_icar = fila_datos["ID SOLICITUD"];
                fila_i.Add("Id.I-Car", id_icar);
                
                //Tipo operacion
                fila_i.Add("Tipo Operación", "AUTOMATICO");

                //Placa patente, vacio hasta hacer coneccion a BD
                fila_i.Add("Placa Patente", "");

                #endregion

                //Añadimos la fila al diccionario que corresponda, dependiendo si es o no contabilizable

                if (contabilizar)
                {
                    _output_datos_a_contabilizar.Add(contador_filas_contabilizadas, fila_i);
                    contador_filas_contabilizadas++;
                }
                else
                {
                    _output_datos_con_problemas_para_contabilizar.Add(contador_filas_no_contabilizadas, fila_i);
                    contador_filas_no_contabilizadas++;
                }
            }

            //Se contabilizan los pagos
            foreach(Pago mi_pago in _mis_pagos)
            {
                
                    Dictionary<string, string> fila_i = new Dictionary<string, string>();
                    #region Llenado de datos
                    //Linea
                    fila_i.Add("Linea", "");
                    // Cuenta
                    string cuenta = "111204";
                    fila_i.Add("Cuenta", cuenta);

                    //Desc. cuenta
                    fila_i.Add("Desc.Cuenta", "");

                    // Debe
                    fila_i.Add("Debe", "0");

                    //Haber
                    fila_i.Add("Haber", mi_pago._monto_pago.ToString());

                    //Glosa
                    string glosa = mi_pago._fecha_pago + " Pago: " + mi_pago._id_pago;
                    fila_i.Add("Glosa", glosa);

                    //Tipo Entidad
                    fila_i.Add("Tipo Entidad", "");

                    //Entidad y nombre entidad
                    fila_i.Add("Entidad", "");
                    fila_i.Add("Nombre Entidad", "");

                    //Tipo Documento
                    string tipo_documento = "TRANSFERENCIA";
                    fila_i.Add("Tipo Docto", tipo_documento);

                    //Emisor Documento
                    fila_i.Add("Emisor Docto", "");

                    //Numero documento
                    string numero_documento = "";
                    fila_i.Add("Nro Docto", mi_pago._id_pago.ToString());

                    //Correlativo documento
                    fila_i.Add("Correlativo Docto", "");

                    //Fecha
                    fila_i.Add("Fecha Vcto", mi_pago._fecha_pago);

                    //Tipo linea
                    fila_i.Add("Tipo Linea", "");

                    //Centro de costo
                    fila_i.Add("Centro de Costo", "");

                    //Sucursal
                    fila_i.Add("Sucursal", "");

                    //N° Doc.Pagado
                    fila_i.Add("N° Doc.Pagado", "");

                    //Folio de factura, vacio
                    fila_i.Add("Folio Fact.", "");

                    //Rut comprador
                    string rut_comprador = "";
                    fila_i.Add("Rut Comprador", rut_comprador);

                    //Rut municipalidad
                    fila_i.Add("Rut Municipalidad", "");

                    //Id i-CAR
                    string id_icar = "";
                    fila_i.Add("Id.I-Car", id_icar);

                    //Tipo operacion
                    fila_i.Add("Tipo Operación", "AUTOMATICO");

                    //Placa patente, vacio hasta hacer coneccion a BD
                    fila_i.Add("Placa Patente", "");
                
                #endregion

                if (mi_pago._pago_es_contabilizable)
                {
                    _output_datos_a_contabilizar.Add(_output_datos_a_contabilizar.Keys.Count, fila_i);                    
                }
                else
                {
                    _output_datos_con_problemas_para_contabilizar.Add(_output_datos_con_problemas_para_contabilizar.Keys.Count, fila_i);
                }
            }

            output_ExcelTabla_a_contabilizar._datos_tabla = _output_datos_a_contabilizar;
            output_ExcelTabla_con_problemas_para_contabilizar._datos_tabla = _output_datos_con_problemas_para_contabilizar;

        }

        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }

        public string convertidor_fecha(string fecha)
        {
            string output = "";
            string[] datos = fecha.Split('-');
            if (datos.Length == 3)
            {
                if (datos[1] == "ene.")
                {
                    datos[1] = "01";
                }
                else if (datos[1] == "feb.")
                {
                    datos[1] = "02";
                }
                else if (datos[1] == "mar.")
                {
                    datos[1] = "03";
                }
                else if (datos[1] == "abr.")
                {
                    datos[1] = "04";
                }
                else if (datos[1] == "may.")
                {
                    datos[1] = "05";
                }
                else if (datos[1] == "jun.")
                {
                    datos[1] = "06";
                }
                else if (datos[1] == "jul.")
                {
                    datos[1] = "07";
                }
                else if (datos[1] == "ago.")
                {
                    datos[1] = "08";
                }
                else if (datos[1] == "sept.")
                {
                    datos[1] = "09";
                }
                else if (datos[1] == "oct.")
                {
                    datos[1] = "10";
                }
                else if (datos[1] == "nov.")
                {
                    datos[1] = "11";
                }
                else if (datos[1] == "dic.")
                {
                    datos[1] = "12";
                }
                else
                {
                    datos[1] = "error en la fecha?";
                }

                output += datos[0] + "-" + datos[1] + "-" + datos[2];
            }

            return output;
        }
    }
}
