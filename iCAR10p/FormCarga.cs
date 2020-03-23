using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;


namespace iCAR10p
{
    public partial class FormCarga : Form
    {
        public FormCarga()
        {
            InitializeComponent();
        }

        // Attemps to read workbook as XLSX, then XLS, then fails.
        public IWorkbook Read_Workbook(string path)
        {
            IWorkbook book = null;

            try
            {
                FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                // Try to read workbook as XLSX:
                try
                {
                    book = new XSSFWorkbook(fs);
                }
                catch
                {
                    book = null;
                }

                // If reading fails, try to read workbook as XLS:
                if (book == null)
                {
                    book = new HSSFWorkbook(fs);
                }

                
            }
            catch (Exception ex)
            { }
            //    MessageBox.Show(ex.Message, "Excel read error", MessageBoxButton.OK, MessageBoxImage.Error);
            //    this.DialogResult = false;
            //    this.Close();

            //}
            return book;
        }

        private void Carga_Archivo_Prendas(object sender, EventArgs e)
        {
            IWorkbook mi_Workbook;
            ISheet mi_Sheet;
            Dictionary<string, TablaExcel> mis_TablasExcel = new Dictionary<string, TablaExcel>();
            //Stream mio;                                                               //En esta parte del codigo se abre un dialogo con
            OpenFileDialog dialogo = new OpenFileDialog();                              // el usuario, y se toma la informacion del archivo,
            dialogo.InitialDirectory = "C:\\Users\\Maximiliano Maroto\\Google Drive\\Contabilidad ICAR\\Prendas y Alzamientos\\Febrero";      // que en caso de no ser correcto lanzara una excepcion
            //dialogo.Filter = "txt files (*.txt)|*.txt";                               // que será atrapada y finalizara el programa.
            dialogo.RestoreDirectory = true;
            dialogo.ShowDialog();

            string path = dialogo.FileName;
            mi_Workbook = this.Read_Workbook(path);

            mi_Sheet = mi_Workbook.GetSheetAt(0);

            string directorio = System.IO.Path.GetDirectoryName(path);


            if(mi_Sheet != null)
            {
                int numero_filas_hoja = mi_Sheet.LastRowNum;
                int fila_inicio_informacion = 0;
                int columna_inicio_informacion = 0;
                int numero_de_columnas = 0;
                string nombre_tabla = "Datos_tramites";
                mis_TablasExcel.Add(nombre_tabla, new TablaExcel(nombre_tabla));

                for (int i = 0; i < numero_filas_hoja; i++)
                {
                    IRow mi_Fila = mi_Sheet.GetRow(i);
                    if (mi_Fila != null)
                    {
                        try
                        {
                            numero_de_columnas = mi_Fila.Cells.Count();
                            fila_inicio_informacion = i;
                            break;
                        }
                        catch (Exception exce)
                        {

                        }
                    }
                }

                for(int c = 0; c < 100; c++) //Buscamos la columna donde comienza la información
                {
                    IRow mi_Fila = mi_Sheet.GetRow(fila_inicio_informacion);
                    if (mi_Fila.GetCell(c)!= null)
                    {
                        columna_inicio_informacion = c;
                        break;
                    }
                }

                for(int k = fila_inicio_informacion; k <= numero_filas_hoja; k++) //Se recorre la tabla que contiene la información
                {
                    IRow mi_Fila = mi_Sheet.GetRow(k);
                    mis_TablasExcel[nombre_tabla]._datos_tabla.Add(k - fila_inicio_informacion, new Dictionary<string, string>());

                    for (int j = columna_inicio_informacion; j < numero_de_columnas + columna_inicio_informacion; j++)
                    {
                        #region Se nombran las columnas
                        if (k == fila_inicio_informacion) //Se nombran las columnas, leyendo la primera fila que trae información
                        {
                            if (mi_Fila.GetCell(j) != null) //Si la celda contine información la obtenemos y agregamos como título
                            {
                                mis_TablasExcel[nombre_tabla]._nombre_columnas_tabla.Add(j - columna_inicio_informacion, mi_Fila.GetCell(j).ToString());
                            }
                            else //En caso contratio ponemos como título Columna con el número de la columna
                            {
                                mis_TablasExcel[nombre_tabla]._nombre_columnas_tabla.Add(j, "Columna " + (j - columna_inicio_informacion).ToString());
                            }
                        }
                        #endregion

                        if (mi_Fila.GetCell(j) != null) //Si la celda contine información la obtenemos y la agregamos
                        {
                            try //Revisamos si es un string o una formula y agregamos el dato
                            {
                                string Campo_RichString = mi_Fila.GetCell(j).RichStringCellValue.ToString();
                                mis_TablasExcel[nombre_tabla]._datos_tabla[k - fila_inicio_informacion].Add(mis_TablasExcel[nombre_tabla]._nombre_columnas_tabla[j - columna_inicio_informacion], Campo_RichString);
                            }
                            catch(Exception f)
                            {
                                int re = 1;
                                mis_TablasExcel[nombre_tabla]._datos_tabla[k - fila_inicio_informacion].Add(mis_TablasExcel[nombre_tabla]._nombre_columnas_tabla[j - columna_inicio_informacion], mi_Fila.GetCell(j).ToString());
                            }
                            
                           
                        }
                        else //En caso contratio la dejamos como un string vacio
                        {
                            mis_TablasExcel[nombre_tabla]._datos_tabla[k - fila_inicio_informacion].Add(mis_TablasExcel[nombre_tabla]._nombre_columnas_tabla[j - columna_inicio_informacion], "");
                        }

                    }

                }

                Logica_Prendas_Alzamientos exportar = new Logica_Prendas_Alzamientos(mis_TablasExcel[nombre_tabla]);
                exportar._directorio = directorio;
                exportar.Procesar_datos();

                Exportar_Archivo(exportar.output_ExcelTabla_a_contabilizar);
                Exportar_Archivo(exportar.output_ExcelTabla_con_problemas_para_contabilizar);

                this.Dispose();
            }
            
        }

        private void Exportar_Archivo(TablaExcel mi_tabla)
        {
            System.Data.DataTable data = new System.Data.DataTable();

            foreach(int mi_indice in mi_tabla._nombre_columnas_tabla.Keys) //Leemos las columnas de la tabla y asignamos el nombre
            {
                string dato = mi_tabla._nombre_columnas_tabla[mi_indice].ToString();
                data.Columns.Add(dato);
            }

            for(int i = 1; i < mi_tabla._datos_tabla.Keys.Count(); i++) //Ubicamos los valores en la tabla de datos
            {
                DataRow nueva_fila = data.NewRow();

                foreach(int mi_indice in mi_tabla._nombre_columnas_tabla.Keys)
                {
                    string dato = mi_tabla._datos_tabla[i][mi_tabla._nombre_columnas_tabla[mi_indice]].ToString();
                    nueva_fila[mi_tabla._nombre_columnas_tabla[mi_indice]] = dato;
                }

                data.Rows.Add(nueva_fila);
            }

            data.TableName = mi_tabla._nombre_tabla;

           DataTable_To_Excel(data, mi_tabla._directorio+"\\" + mi_tabla._nombre_tabla+".xlsx");

        }

        /// <summary>Convierte un DataTable en un archivo de Excel (xls o Xlsx) y lo guarda en disco.</summary>
        /// <param name="pDatos">Datos de la Tabla a guardar. Usa el nombre de la tabla como nombre de la Hoja</param>
        /// <param name="pFilePath">Ruta del archivo donde se guarda.</param>
        private void DataTable_To_Excel(System.Data.DataTable pDatos, string pFilePath)
        {
            try
            {
                if (pDatos != null && pDatos.Rows.Count > 0)
                {
                    IWorkbook workbook = null;
                    ISheet worksheet = null;

                    using (FileStream stream = new FileStream(pFilePath, FileMode.Create, FileAccess.ReadWrite))
                    {
                        string Ext = System.IO.Path.GetExtension(pFilePath); //<-Extension del archivo
                        switch (Ext.ToLower())
                        {
                            case ".xls":
                                HSSFWorkbook workbookH = new HSSFWorkbook();
                                NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
                                dsi.Company = "Cutcsa"; dsi.Manager = "Departamento Informatico";
                                workbookH.DocumentSummaryInformation = dsi;
                                workbook = workbookH;
                                break;

                            case ".xlsx": workbook = new XSSFWorkbook(); break;
                        }

                        worksheet = workbook.CreateSheet(pDatos.TableName); //<-Usa el nombre de la tabla como nombre de la Hoja

                        //CREAR EN LA PRIMERA FILA LOS TITULOS DE LAS COLUMNAS
                        int iRow = 0;
                        if (pDatos.Columns.Count > 0)
                        {
                            int iCol = 0;
                            IRow fila = worksheet.CreateRow(iRow);
                            foreach (DataColumn columna in pDatos.Columns)
                            {
                                ICell cell = fila.CreateCell(iCol, CellType.String);
                                cell.SetCellValue(columna.ColumnName);
                                iCol++;
                            }
                            iRow++;
                        }

                        //FORMATOS PARA CIERTOS TIPOS DE DATOS
                        ICellStyle _doubleCellStyle = workbook.CreateCellStyle();
                        _doubleCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.###");

                        ICellStyle _intCellStyle = workbook.CreateCellStyle();
                        _intCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");

                        ICellStyle _boolCellStyle = workbook.CreateCellStyle();
                        _boolCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("BOOLEAN");

                        ICellStyle _dateCellStyle = workbook.CreateCellStyle();
                        _dateCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy");

                        ICellStyle _dateTimeCellStyle = workbook.CreateCellStyle();
                        _dateTimeCellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("dd-MM-yyyy HH:mm:ss");

                        //AHORA CREAR UNA FILA POR CADA REGISTRO DE LA TABLA
                        foreach (DataRow row in pDatos.Rows)
                        {
                            IRow fila = worksheet.CreateRow(iRow);
                            int iCol = 0;
                            foreach (DataColumn column in pDatos.Columns)
                            {
                                ICell cell = null; //<-Representa la celda actual                               
                                object cellValue = row[iCol]; //<- El valor actual de la celda

                                switch (column.DataType.ToString())
                                {
                                    case "System.Boolean":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Boolean);

                                            if (Convert.ToBoolean(cellValue)) { cell.SetCellFormula("TRUE()"); }
                                            else { cell.SetCellFormula("FALSE()"); }

                                            cell.CellStyle = _boolCellStyle;
                                        }
                                        break;

                                    case "System.String":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.String);
                                            cell.SetCellValue(Convert.ToString(cellValue));
                                        }
                                        break;

                                    case "System.Int32":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt32(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Int64":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToInt64(cellValue));
                                            cell.CellStyle = _intCellStyle;
                                        }
                                        break;
                                    case "System.Decimal":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;
                                    case "System.Double":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDouble(cellValue));
                                            cell.CellStyle = _doubleCellStyle;
                                        }
                                        break;

                                    case "System.DateTime":
                                        if (cellValue != DBNull.Value)
                                        {
                                            cell = fila.CreateCell(iCol, CellType.Numeric);
                                            cell.SetCellValue(Convert.ToDateTime(cellValue));

                                            //Si No tiene valor de Hora, usar formato dd-MM-yyyy
                                            DateTime cDate = Convert.ToDateTime(cellValue);
                                            if (cDate != null && cDate.Hour > 0) { cell.CellStyle = _dateTimeCellStyle; }
                                            else { cell.CellStyle = _dateCellStyle; }
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                iCol++;
                            }
                            iRow++;
                        }

                        workbook.Write(stream);
                        stream.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Generar_pdf_tag(object sender, EventArgs e)
        {
            IWorkbook mi_Workbook;
            ISheet mi_Sheet;
            List<string> placas_patentes_a_generar_tag = new List<string>();
            //Stream mio;                                                               //En esta parte del codigo se abre un dialogo con
            OpenFileDialog dialogo = new OpenFileDialog();                              // el usuario, y se toma la informacion del archivo,
            dialogo.InitialDirectory = "C:\\Users\\Maximiliano Maroto\\Documents";      // que en caso de no ser correcto lanzara una excepcion
            //dialogo.Filter = "txt files (*.txt)|*.txt";                               // que será atrapada y finalizara el programa.
            dialogo.RestoreDirectory = true;
            dialogo.ShowDialog();

            string path = dialogo.FileName;
            mi_Workbook = this.Read_Workbook(path);
            mi_Sheet = mi_Workbook.GetSheetAt(0);


            if (mi_Sheet != null)
            {
                int numero_filas_hoja = mi_Sheet.LastRowNum;
                int fila_inicio_informacion = 0;
                int columna_inicio_informacion = 0;
                int numero_de_columnas = 0;


                for (int i = 0; i < numero_filas_hoja; i++)
                {
                    IRow mi_Fila = mi_Sheet.GetRow(i);
                    if (mi_Fila != null)
                    {
                        try
                        {
                            numero_de_columnas = mi_Fila.Cells.Count();
                            fila_inicio_informacion = i;
                            break;
                        }
                        catch (Exception exce)
                        {

                        }
                    }
                }

                for (int c = 0; c < 100; c++) //Buscamos la columna donde comienza la información
                {
                    IRow mi_Fila = mi_Sheet.GetRow(fila_inicio_informacion);
                    if (mi_Fila.GetCell(c) != null)
                    {
                        columna_inicio_informacion = c;
                        break;
                    }
                }

                for (int k = fila_inicio_informacion; k <= numero_filas_hoja; k++) //Se recorre la tabla que contiene la información
                {
                    IRow mi_Fila = mi_Sheet.GetRow(k);


                    for (int j = columna_inicio_informacion; j < numero_de_columnas + columna_inicio_informacion; j++)
                    {

                        if (mi_Fila.GetCell(j) != null) //Si la celda contine información la obtenemos y agregamos como título
                        {
                            if (!placas_patentes_a_generar_tag.Contains(mi_Fila.GetCell(j).ToString()))
                            {
                                placas_patentes_a_generar_tag.Add(mi_Fila.GetCell(j).ToString());
                            }
                        }

                    }

                }

                foreach(string mi_placa in placas_patentes_a_generar_tag)
                {
                    File.Copy(@"C:\\Users\\Maximiliano Maroto\\source\\repos\\iCAR10p\\Archivos\\blanco.pdf", @"C:\\Users\\Maximiliano Maroto\\source\\repos\\iCAR10p\\Archivos\\tag_"+mi_placa+".pdf");
                }

                this.Dispose();
            }
        }

        private void Leer_facturas_iCAR_obtener_patentes_facturadas(object sender, EventArgs e)
        {
            string nombre_tabla = "Patentes facturadas";
            TablaExcel patentes_facturadas = new TablaExcel(nombre_tabla);

            OpenFileDialog dialogo = new OpenFileDialog();                              
            dialogo.InitialDirectory = "C:\\Users\\Maximiliano Maroto\\Google Drive\\Contabilidad ICAR\\Facturas 2020";      
            dialogo.RestoreDirectory = true;
            dialogo.ShowDialog();
            string path = dialogo.FileName;
            string directorio = System.IO.Path.GetDirectoryName(path);

            Logica_lectura_facturas Lector = new Logica_lectura_facturas(directorio);
            Lector.Procesar_facturas();
            Lector.Generar_Tablas_Excel();

            Exportar_Archivo(Lector._output_patentes);

            this.Dispose();
        }
    }
}
