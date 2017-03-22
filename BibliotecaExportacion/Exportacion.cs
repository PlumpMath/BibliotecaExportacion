using System;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel.DataAnnotations;
using System.Resources;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportLibrary
{
    /// <summary>
    /// Clase que contiene métodos estáticos para la exportación de coleccion de datos a formatos específicos
    /// </summary>
    public class Export
    {
        

        /// <summary>
        /// Constructor de la clase que carga las dependencias 
        /// </summary>
        public Export()
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
        }

        System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string dllName = args.Name.Contains(',') ? args.Name.Substring(0, args.Name.IndexOf(',')) : args.Name.Replace(".dll", "");

            dllName = dllName.Replace(".", "_");

            if (dllName.EndsWith("_resources")) return null;

            System.Resources.ResourceManager rm = new System.Resources.ResourceManager(GetType().Namespace + ".Properties.Resources", System.Reflection.Assembly.GetExecutingAssembly());

            byte[] bytes = (byte[])rm.GetObject(dllName);

            return System.Reflection.Assembly.Load(bytes);
        }

        /// <summary>
        /// Método para exportar a formato Excel una colección de objetos a partir de Interop Excel.
        /// </summary>
        /// <typeparam name="T">Clase que define a la colección de objetos</typeparam>
        /// <param name="_data">Colección de objetos tipo T a exportar</param>
        /// <param name="_path">Path de almacenamiento temporal</param>
        /// <param name="_printHeader">Indica si se imprime los títulos en la cabecera</param>
        /// <returns>Regresa un Tuple con el archivo generado en el primer item 
        /// y un comentario en el segundo item</returns>
        /// <remarks>Este método hace uso de la interoperabilidad de Excel.
        /// El nombre de la columna se toma del atributo DisplatAttribute </remarks>
        public Tuple<byte[], string> ExportToExcel<T>(IEnumerable<T> _data, string _path,
            bool _printHeader = true) where T : class
        {
            try
            {
                if (_data == null || !_data.Any())
                    return new Tuple<byte[], string>(null, Resources.DataEmpty);
                if (string.IsNullOrEmpty(_path))
                    return new Tuple<byte[], string>(null, Resources.PathEmpty);
                PropertyInfo[] _properties = typeof(T).GetProperties();
                Microsoft.Office.Interop.Excel.Application _excel = new Microsoft.Office.Interop.Excel.Application();
                _excel.Workbooks.Add();
                Microsoft.Office.Interop.Excel._Worksheet _workSheet =
                    (Microsoft.Office.Interop.Excel._Worksheet) _excel.ActiveSheet;
                if (_printHeader)
                {
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        var _displayAttribute =
                            _properties[_counter].GetCustomAttributes(false)
                                .FirstOrDefault(a => a is DisplayAttribute) as DisplayAttribute;
                        _workSheet.Cells[1, _counter + 1] = _displayAttribute != null
                            ? GetDisplayAttributeFromResourceType(_displayAttribute)
                            : _properties[_counter].Name;
                    }
                }
                int _counter2 = 2;
                foreach (var _element in _data)
                {
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        if (_properties[_counter].GetValue(_element) != null)
                        {
                            switch (_properties[_counter].GetValue(_element).GetType().ToString())
                            {
                                case "System.String":
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                case "System.Decimal":
                                case "System.Double":
                                case "System.DateTime":
                                    _workSheet.Cells[_counter2, _counter + 1] =
                                        _properties[_counter].GetValue(_element);
                                    break;
                                default:
                                    _workSheet.Cells[_counter2, _counter + 1] = Resources.DefaultTypeText;
                                    break;
                            }
                        }
                        else
                            _workSheet.Cells[_counter2, _counter + 1] = string.Empty;
                    }
                    _counter2++;
                }
                _workSheet.Range["A1"].AutoFormat();
                _workSheet.SaveAs(_path);
                byte[] _file = System.IO.File.ReadAllBytes(_path);
                _excel.Quit();
                if (_excel != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_excel);
                if (_workSheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_workSheet);
                _excel = null;
                _workSheet = null;
                GC.Collect();
                System.IO.File.Delete(_path);
                return new Tuple<byte[], string>(_file, null);
            }
            catch (Exception ex)
            {
                return new Tuple<byte[], string>(null, Resources.ExceptionText + ex.Message);
            }
        }

        /// <summary>
        /// Método para obtener el valor del DisplayAttribute desde el archivo de recursos
        /// </summary>
        /// <param name="_displayAttribute">Objeto que identifica a la propiedad actual</param>
        /// <returns>Retorna el valor que tiene en el archivo de recursos</returns>
        private string GetDisplayAttributeFromResourceType(DisplayAttribute _displayAttribute)
        {
            try
            {
                ResourceManager resourceManager = new ResourceManager(_displayAttribute.ResourceType);
                return resourceManager.GetResourceSet(Thread.CurrentThread.CurrentUICulture, true, true)
                                            .OfType<DictionaryEntry>()
                                            .FirstOrDefault(p => p.Key.ToString() == _displayAttribute.Name).Value.ToString();
            }
            catch (Exception)
            {
                return _displayAttribute.Name;
            }
        }

        /// <summary>
        /// Método para exportar a formato Excel una colección de objetos, a partir de un XML (spreedsheet).
        /// </summary>
        /// <typeparam name="T">Clase que define a la colección de objetos</typeparam>
        /// <param name="_data">Colección de objetos tipo T a exportar</param>
        /// <param name="_columnsToPrint">Arreglo de cadenas con los nombres de las propiedades a imprimir</param>
        /// <param name="_dateTimeFormat">Representa el formato de impresión para propiedades del objeto tipo DateTime</param>
        /// <param name="_printHeader">Indica si se imprime los títulos en la cabecera</param>
        /// <returns>Regresa un Tuple con el archivo generado en el primer item 
        /// y un comentario en el segundo item</returns>
        /// <remarks>El nombre de la columna se toma del atributo DisplatAttribute</remarks>
        public Tuple<byte[], string> ExportToExcelXml<T>(IEnumerable<T> _data, IEnumerable<string> _columnsToPrint = null, string _dateTimeFormat = "dd/MM/yyyy", bool _printHeader = true) where T : class
        {
            try
            {
                if (_data == null || !_data.Any())
                    return new Tuple<byte[], string>(null, Resources.DataEmpty);
                PropertyInfo[] _properties = typeof(T).GetProperties();
                System.IO.MemoryStream _stream = new System.IO.MemoryStream();
                System.IO.StreamWriter _excelDoc = new System.IO.StreamWriter(_stream);
                _excelDoc.Write(ResourcesXMLExcel.HeaderXML);
                _excelDoc.Write(ResourcesXMLExcel.OpenWorkSheet);
                _excelDoc.Write(ResourcesXMLExcel.OpenTable);
                bool _print;
                if (_printHeader)
                {
                    _excelDoc.Write(ResourcesXMLExcel.OpenRowWithStyle);
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(x => x.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            _excelDoc.Write(ResourcesXMLExcel.OpenCellString);
                            var _displayAttribute =
                                _properties[_counter].GetCustomAttributes(false)
                                    .FirstOrDefault(a => a is DisplayAttribute) as DisplayAttribute;
                            
                            _excelDoc.Write(_displayAttribute != null
                                ? GetDisplayAttributeFromResourceType(_displayAttribute)
                                : _properties[_counter].Name);
                            _excelDoc.Write(ResourcesXMLExcel.CloseCell); 
                        }
                    }
                    _excelDoc.Write(ResourcesXMLExcel.CloseRow); 
                }
                foreach (var x in _data)
                {
                    _excelDoc.Write(ResourcesXMLExcel.OpenRow);
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(y => y.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            var _value = _properties[_counter].GetValue(x);
                            switch (_value != null ? _value.GetType().ToString() : string.Empty)
                            {
                                case "":
                                case "System.String":
                                    string XMLstring = _value != null ? _value.ToString() : string.Empty;
                                    XMLstring = XMLstring.Trim();
                                    XMLstring = XMLstring.Replace("&", "");
                                    XMLstring = XMLstring.Replace(">", "");
                                    XMLstring = XMLstring.Replace("<", "");
                                    _excelDoc.Write(ResourcesXMLExcel.OpenCellString);
                                    _excelDoc.Write(XMLstring);
                                    _excelDoc.Write(ResourcesXMLExcel.CloseCell);
                                    break;
                                case "System.DateTime":
                                    _excelDoc.Write(ResourcesXMLExcel.OpenCellString);
                                    _excelDoc.Write(_value != null
                                        ? ((DateTime) _value).ToString(_dateTimeFormat)
                                        : string.Empty);
                                    _excelDoc.Write(ResourcesXMLExcel.CloseCell);
                                    break;
                                case "System.Boolean":
                                    _excelDoc.Write(ResourcesXMLExcel.OpenCellString);
                                    _excelDoc.Write(_value != null ? ((bool) _value ? ResourcesXMLExcel.TrueValue : ResourcesXMLExcel.FalseValue) : string.Empty);
                                    _excelDoc.Write(ResourcesXMLExcel.CloseCell);
                                    break;
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                case "System.Decimal":
                                case "System.Double":
                                    _excelDoc.Write(ResourcesXMLExcel.OpenCellNumber);
                                    _excelDoc.Write(_value != null ? _value.ToString() : string.Empty);
                                    _excelDoc.Write(ResourcesXMLExcel.CloseCell);
                                    break;
                                case "System.DBNull":
                                    _excelDoc.Write(ResourcesXMLExcel.OpenCellString);
                                    _excelDoc.Write(Resources.NullObject);
                                    _excelDoc.Write(ResourcesXMLExcel.CloseCell);
                                    break;
                                default:
                                    _excelDoc.Write(ResourcesXMLExcel.OpenCellString);
                                    _excelDoc.Write(Resources.DefaultTypeText);
                                    _excelDoc.Write(ResourcesXMLExcel.CloseCell);
                                    break;
                            }
                        }
                    }
                    _excelDoc.Write(ResourcesXMLExcel.CloseRow); 
                }
                _excelDoc.Write(ResourcesXMLExcel.CloseTable);
                _excelDoc.Write(ResourcesXMLExcel.CloseWorkSheet);
                _excelDoc.Write(ResourcesXMLExcel.FooterXML);
                _excelDoc.Close();
                return new Tuple<byte[], string>(_stream.ToArray(), null);
            }
            catch (Exception ex)
            {
                return new Tuple<byte[], string>(null, Resources.ExceptionText + ex.Message);
            }
        }

        /// <summary>
        /// Indica el tamaño de la celda
        /// </summary>
        /// <param name="intCol">Entero que indica la columna</param>
        /// <returns></returns>
        private string ColumnLetter(int intCol)
        {
            var intFirstLetter = ((intCol) / 676) + 64;
            var intSecondLetter = ((intCol % 676) / 26) + 64;
            var intThirdLetter = (intCol % 26) + 65;

            var firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }

        /// <summary>
        /// Método para crear una celda con texto
        /// </summary>
        /// <param name="header">Encabezado</param>
        /// <param name="index">Entero que indica posición</param>
        /// <param name="text">Texto de la celda</param>
        /// <param name="indexFont">Indice del estilo aplicable a la celda</param>
        /// <returns>Objeto de tipo Cell</returns>
        private Cell CreateTextCell(string header, uint index, string text, uint? indexFont = null)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = header + index,
            };
            if(indexFont!= null)
                cell.StyleIndex = (uint)indexFont;
            InlineString istring = new InlineString();
            Text t = new Text { Text = text };
            istring.AppendChild(t);
            cell.AppendChild(istring);
            return cell;
        }

        /// <summary>
        /// Método de agrega una hoja de estilos
        /// </summary>
        /// <param name="spreadsheet">Documento original</param>
        /// <returns></returns>
        private WorkbookStylesPart AddStyleSheet(SpreadsheetDocument spreadsheet)
        {
            WorkbookStylesPart stylesheet = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet workbookstylesheet = new Stylesheet();
            DocumentFormat.OpenXml.Spreadsheet.Font font0 = new DocumentFormat.OpenXml.Spreadsheet.Font();         // Default font
            DocumentFormat.OpenXml.Spreadsheet.Font font1 = new DocumentFormat.OpenXml.Spreadsheet.Font();         // Bold font
            Bold bold = new Bold();
            font1.Append(bold);
            Fonts fonts = new Fonts();      // <APENDING Fonts>
            fonts.Append(font0);
            fonts.Append(font1);
            // <Fills>
            Fill fill0 = new Fill();        // Default fill
            Fills fills = new Fills();      // <APENDING Fills>
            fills.Append(fill0);
            // <Borders>
            Border border0 = new Border();     // Defualt border
            Borders borders = new Borders();    // <APENDING Borders>
            borders.Append(border0);
            // <CellFormats>
            CellFormat cellformat0 = new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }; // Default style : Mandatory | Style ID =0
            CellFormat cellformat1 = new CellFormat() { FontId = 1 }; // Style with Bold text ; Style ID = 1
            // <APENDING CellFormats>
            CellFormats cellformats = new CellFormats();
            cellformats.Append(cellformat0);
            cellformats.Append(cellformat1);
            // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet <Preserve the ORDER>
            workbookstylesheet.Append(fonts);
            workbookstylesheet.Append(fills);
            workbookstylesheet.Append(borders);
            workbookstylesheet.Append(cellformats);
            // Finalize
            stylesheet.Stylesheet = workbookstylesheet;
            stylesheet.Stylesheet.Save();
            return stylesheet;
        }

        /// <summary>
        /// Método para exportar a formato Excel una colección de objetos, a partir de un XML OpenXML.
        /// </summary>
        /// <typeparam name="T">Clase que define a la colección de objetos</typeparam>
        /// <param name="_data">Colección de objetos tipo T a exportar</param>
        /// <param name="_columnsToPrint">Arreglo de cadenas con los nombres de las propiedades a imprimir</param>
        /// <param name="_dateTimeFormat">Representa el formato de impresión para propiedades del objeto tipo DateTime</param>
        /// <param name="_printHeader">Indica si se imprime los títulos en la cabecera</param>
        /// <param name="_sheetName">Nombre de la hoja activa del libro</param>
        /// <returns>Regresa un Tuple con el archivo generado en el primer item 
        /// y un comentario en el segundo item</returns>
        /// <remarks>El nombre de la columna se toma del atributo DisplatAttribute</remarks>
        public Tuple<byte[], string> ExportToExcelOpenXml<T>(IEnumerable<T> _data, IEnumerable<string> _columnsToPrint = null, string _dateTimeFormat = "dd/MM/yyyy", bool _printHeader = true, string _sheetName = null) where T : class
        {
            try
            {
                if (_data == null || !_data.Any())
                    return new Tuple<byte[], string>(null, String.Empty);
                System.IO.MemoryStream stream = new System.IO.MemoryStream();
                SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookpart = document.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = _sheetName ?? Resources.DefaultSheetName
                };
                sheets.AppendChild(sheet);
                AddStyleSheet(document);
                PropertyInfo[] _properties = typeof(T).GetProperties();
                UInt32 rowIdex = 0;
                var row = new Row { RowIndex = ++rowIdex };
                sheetData.AppendChild(row);
                var cellIdex = 0;
                bool _print;
                if (_printHeader)
                {
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(x => x.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            var _displayAttribute =
                                _properties[_counter].GetCustomAttributes(false)
                                    .FirstOrDefault(a => a is DisplayAttribute) as DisplayAttribute;
                            row.AppendChild(CreateTextCell(ColumnLetter(cellIdex++), rowIdex, _displayAttribute != null
                                ? GetDisplayAttributeFromResourceType(_displayAttribute)
                                : _properties[_counter].Name, 1));
                        }
                    }
                }
                foreach (var x in _data)
                {
                    cellIdex = 0;
                    row = new Row { RowIndex = ++rowIdex };
                    sheetData.AppendChild(row);
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(y => y.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            Cell cell;
                            var _value = _properties[_counter].GetValue(x);
                            switch (_value != null ? _value.GetType().ToString() : string.Empty)
                            {
                                case "":
                                case "System.String":
                                    cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, _value != null ? _value.ToString() : string.Empty);
                                    break;
                                case "System.DateTime":
                                    cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, _value != null
                                        ? ((DateTime)_value).ToString(_dateTimeFormat)
                                        : string.Empty);
                                    break;
                                case "System.Boolean":
                                    cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, _value != null ? ((bool)_value ? ResourcesXMLExcel.TrueValue : ResourcesXMLExcel.FalseValue) : string.Empty);
                                    break;
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                case "System.Decimal":
                                case "System.Double":
                                    cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, _value != null ? _value.ToString() : string.Empty);
                                    break;
                                case "System.DBNull":
                                    cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, Resources.NullObject);
                                    break;
                                default:
                                    cell = CreateTextCell(ColumnLetter(cellIdex++), rowIdex, Resources.DefaultTypeText);
                                    break;
                            }
                            row.AppendChild(cell);
                        }
                    }
                }
                workbookpart.Workbook.Save();
                document.Close();
                return new Tuple<byte[], string>(stream.ToArray(), null);
            }
            catch (Exception ex)
            {
                return new Tuple<byte[], string>(null, Resources.ExceptionText + ex.Message);
            }
        }

        /// <summary>
        /// Método para exportar a formato CSV una colección de objetos
        /// </summary>
        /// <typeparam name="T">Clase que define a la colección de objetos</typeparam>
        /// <param name="_data">Colección de objetos tipo T a exportar</param>
        /// <param name="_separator">Carácter que indica la separación de columnas</param>
        /// <param name="_columnsToPrint">Arreglo de cadenas con los nombres de las propiedades a imprimir</param>
        /// <param name="_dateTimeFormat">Representa el formato de impresión para propiedades del objeto tipo DateTime</param>
        /// <param name="_printHeader">Indica si se imprime los títulos en la cabecera</param>
        /// <returns>Regresa un Tuple con el archivo generado en el primer item 
        /// y un comentario en el segundo item</returns>
        /// <remarks>El nombre de la columna se toma del atributo DisplatAttribute</remarks>
        public Tuple<byte[], string> ExportToCsv<T>(IEnumerable<T> _data, string _separator = ",", IEnumerable<string> _columnsToPrint = null, string _dateTimeFormat = "dd/MM/yyyy", bool _printHeader = true) where T : class
        {
            try
            {
                if (_data == null || !_data.Any())
                    return new Tuple<byte[], string>(null, Resources.DataEmpty);
                System.Text.StringBuilder _result = new System.Text.StringBuilder();
                PropertyInfo[] _properties = typeof(T).GetProperties();
                System.Text.StringBuilder _temporal = new System.Text.StringBuilder();
                bool _print;
                if (_printHeader)
                {
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(x => x.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            var _displayAttribute =
                                _properties[_counter].GetCustomAttributes(false)
                                    .FirstOrDefault(a => a is DisplayAttribute) as DisplayAttribute;
                            _temporal.Append(string.Format("{0}" + _separator,
                                _displayAttribute != null
                                    ? GetDisplayAttributeFromResourceType(_displayAttribute).Replace(_separator, "")
                                    : _properties[_counter].Name));
                        }
                    }
                    _result.AppendLine(_temporal.ToString());
                }
                foreach (var _element in _data)
                {
                    _temporal = new System.Text.StringBuilder();
                    for (int i = 0; i < _properties.Length; i++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(x => x.ToLowerInvariant())
                                .Contains(_properties[i].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            var _value = _properties[i].GetValue(_element);
                            switch (_value != null ? _value.GetType().ToString() : string.Empty)
                            {
                                case "System.DateTime":
                                    if (_value != null)
                                    {
                                        DateTime? _temporalDate = (DateTime) _value;
                                        _temporal.Append((_temporalDate != null
                                                             ? ((DateTime) _temporalDate).ToString(_dateTimeFormat)
                                                             : string.Empty) + _separator);
                                    }
                                    break;
                                case "System.Boolean":
                                    bool? _bool = _value != null && (bool) _value;
                                    _temporal.Append((_bool != null && (bool) _bool
                                                         ? ResourcesXMLExcel.TrueValue
                                                         : ResourcesXMLExcel.FalseValue) + _separator);
                                    break;
                                case "":
                                case "System.String":
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                case "System.Decimal":
                                case "System.Double":
                                case "System.DBNull":
                                    _temporal.Append((_value != null
                                                         ? _value.ToString()
                                                             .Replace(",", string.Empty)
                                                             .Replace("\n", " ")
                                                             .Replace(_separator, "")
                                                         : string.Empty) + _separator);
                                    break;
                                default:
                                    _temporal.Append(Resources.DefaultTypeText + _separator);
                                    break;
                            }
                        }
                    }
                    _result.AppendLine(_temporal.ToString());
                }
                System.Text.Encoding encoding = System.Text.Encoding.UTF8;
                return new Tuple<byte[], string>(encoding.GetBytes(_result.ToString()), null);
            }
            catch (Exception ex)
            {
                return new Tuple<byte[], string>(null, Resources.ExceptionText + ex.Message);
            }
        }

        /// <summary>
        /// Método para exportar a formato PDF una colección de objetos
        /// </summary>
        /// <typeparam name="T">Clase que define a la colección de objetos</typeparam>
        /// <param name="_data">Colección de objetos tipo T a exportar</param>
        /// <param name="_typeSheet">Enum que indica el tipo de hoja y orientación del PDF</param>
        /// <param name="_columnsToPrint">Arreglo de cadenas con los nombres de las propiedades a imprimir</param>
        /// <param name="_dateTimeFormat">Representa el formato de impresión para propiedades del objeto tipo DateTime</param>
        /// <param name="_printHeader">Indica si se imprime los títulos en la cabecera</param>
        /// <returns>Regresa un Tuple con el archivo generado en el primer item 
        /// y un comentario en el segundo item</returns>
        /// <remarks>Este método hace uso de la librería iTextSharp.
        /// El nombre de la columna se toma del atributo DisplatAttribute</remarks>
        public Tuple<byte[], string> ExportToPdf<T>(IEnumerable<T> _data, EnumExport.TypeSheet _typeSheet, IEnumerable<string> _columnsToPrint = null, string _dateTimeFormat = "dd/MM/yyyy", bool _printHeader = true) where T : class
        {
            try
            {
                PropertyInfo[] _properties = typeof(T).GetProperties();
                iTextSharp.text.pdf.PdfPTable table =
                    new iTextSharp.text.pdf.PdfPTable(_columnsToPrint != null && _columnsToPrint.Any()
                        ? _columnsToPrint.Count()
                        : _properties.Length) {WidthPercentage = 100f};
                iTextSharp.text.Font font8 = iTextSharp.text.FontFactory.GetFont("ARIAL", 9);
                bool _print;
                if (_printHeader)
                {
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(x => x.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            var _displayAttribute =
                                _properties[_counter].GetCustomAttributes(false)
                                    .FirstOrDefault(a => a is DisplayAttribute) as DisplayAttribute;

                            iTextSharp.text.pdf.PdfPCell cell =
                                new iTextSharp.text.pdf.PdfPCell(
                                    new iTextSharp.text.Phrase(
                                        new iTextSharp.text.Chunk(
                                            _displayAttribute != null
                                                ? GetDisplayAttributeFromResourceType(_displayAttribute)
                                                : _properties[_counter].Name,
                                            font8)));
                            table.AddCell(cell);
                        }
                    }
                }
                font8 = iTextSharp.text.FontFactory.GetFont("ARIAL", 7);
                foreach (var x in _data)
                {
                    for (int _counter = 0; _counter < _properties.Length; _counter++)
                    {
                        _print = true;
                        if (_columnsToPrint != null && _columnsToPrint.Any())
                        {
                            _print = _columnsToPrint.Select(z => z.ToLowerInvariant())
                                .Contains(_properties[_counter].Name.ToLowerInvariant());
                        }
                        if (_print)
                        {
                            iTextSharp.text.pdf.PdfPCell cell = null;
                            var _value = _properties[_counter].GetValue(x);
                            switch (
                                _value != null ? _properties[_counter].GetValue(x).GetType().ToString() : string.Empty)
                            {
                                case "System.DateTime":

                                    if (_value != null)
                                    {
                                        DateTime? _datetime = (DateTime) _value;
                                        cell =
                                            new iTextSharp.text.pdf.PdfPCell(
                                                new iTextSharp.text.Phrase(
                                                    new iTextSharp.text.Chunk(
                                                        ((DateTime) _datetime).ToString(_dateTimeFormat), font8)));
                                    }
                                    break;
                                case "":
                                case "System.String":
                                case "System.Boolean":
                                case "System.Int16":
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                case "System.Decimal":
                                case "System.Double":
                                case "System.DBNull":
                                    cell =
                                        new iTextSharp.text.pdf.PdfPCell(
                                            new iTextSharp.text.Phrase(
                                                new iTextSharp.text.Chunk(
                                                    _value != null ? _value.ToString() : string.Empty, font8)));
                                    break;
                                default:
                                    cell =
                                        new iTextSharp.text.pdf.PdfPCell(
                                            new iTextSharp.text.Phrase(
                                                new iTextSharp.text.Chunk(Resources.DefaultTypeText, font8)));
                                    break;
                            }
                            cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            table.AddCell(cell);
                        }
                    }
                }
                iTextSharp.text.Document pdfDoc;
                switch (_typeSheet)
                {
                    case EnumExport.TypeSheet.Letter:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER, 10, 10, 10, 10);
                        break;
                    case EnumExport.TypeSheet.LetterHorizontal:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), 10, 10, 10, 10);
                        break;
                    case EnumExport.TypeSheet.A4:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 10, 10, 10, 10);
                        break;
                    case EnumExport.TypeSheet.A4Horizontal:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 10, 10, 10, 10);
                        break;
                    case EnumExport.TypeSheet.Legal:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LEGAL, 10, 10, 10, 10);
                        break;
                    case EnumExport.TypeSheet.LegalHorizontal:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LEGAL.Rotate(), 10, 10, 10, 10);
                        break;
                    default:
                        pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER, 10, 10, 10, 10);
                        break;
                }
                System.IO.MemoryStream _stream = new System.IO.MemoryStream();
                iTextSharp.text.pdf.PdfWriter.GetInstance(pdfDoc, _stream);
                pdfDoc.Open();
                pdfDoc.Add(table);
                pdfDoc.Close();
                return new Tuple<byte[], string>(_stream.ToArray(), null);
            }
            catch (Exception ex)
            {
                return new Tuple<byte[], string>(null, Resources.ExceptionText + ex.Message);
            }
        }
        
        //public  void WriteListToExcel<T>(IEnumerable<T> list, string fulllPath) where T : class
        //{
        //    try
        //    {
        //        List<string> result = new List<string>();
        //        result.Add(String.Join(String.Empty, typeof(T).GetProperties().Select(i => String.Format("{0}\t", i.Name)))); // Headers
        //        result.AddRange(list.Select(i => String.Join("\t", i.GetType().GetProperties().Select(t => t.GetValue(i, null)))));
        //        // Lines

        //        File.WriteAllLines(fulllPath, result);
        //    }
        //    catch (Exception e)
        //    {
        //        throw new Exception(e.Message);
        //    }
        //}

        //public  void WriteListToExcel2<T>(HttpResponseBase Response, IEnumerable<T> list, string fileName)
        //    where T : class
        //{
        //    try
        //    {
        //        Response.Clear();
        //        Response.AddHeader("content-disposition", String.Format("attachment;filename={0}.xls", fileName));
        //        Response.Charset = String.Empty;
        //        Response.Cache.SetCacheability(HttpCacheability.NoCache);
        //        Response.ContentType = "application/vnd.ms-excel";

        //        List<string> result = new List<string>();
        //        result.Add(String.Format("{0}\n",
        //                String.Join(String.Empty, typeof(T).GetProperties().Select(i => String.Format("{0}\t", i.Name)))));
        //        // Headers
        //        result.AddRange(
        //                list.Select(
        //                    i =>
        //                        String.Format("{0}\n",
        //                            String.Join("\t", i.GetType().GetProperties().Select(t => t.GetValue(i, null))))));
        //        // Lines 
        //        result.ForEach(i => Response.Write(i));

        //        Response.Flush();
        //        Response.End();
        //    }
        //    catch (Exception e)
        //    {
        //        // Error..
        //    }
        //}

        ///// <summary>
        ///// generate an xml-spreadsheet document from a data-table.
        ///// </summary>
        ///// <param name="tbl"></param>
        ///// <returns></returns>
        //public  XDocument GenerateXMLSpreadSheet<T>(IEnumerable<T> _data) where T : class
        //{
        //    XDocument xmlssDoc =
        //        new XDocument(
        //            new XDeclaration("1.0", "utf-8", "yes"),
        //            new XProcessingInstruction("mso-application", "Excel.Sheet"),
        //            new XElement("Workbook",
        //                new XAttribute("xmlns", "urn:schemas-microsoft-com:office:spreadsheet"),
        //                new XAttribute("xmlns:ss", "urn:schemas-microsoft-com:office:spreadsheet"),
        //                new XElement("Worksheet", new XAttribute("ss:Name", "Coleccion"),
        //                    new XElement("Table", GetRows(_data)
        //                    )
        //                )
        //            )
        //        );
        //    return xmlssDoc;
        //}

        ///// <summary>
        ///// generate xml-spreadsheet rows from a data-table.
        ///// </summary>
        ///// <param name="tbl"></param>
        ///// <returns></returns>
        //public  Object[] GetRows<T>(IEnumerable<T> _data) where T : class
        //{
        //    try
        //    {


        //        PropertyDescriptorCollection _properties = TypeDescriptor.GetProperties(typeof(T));
        //        XElement[] rows = new XElement[_data.Count()];

        //        int r = 0;
        //        foreach (var row in _data)
        //        {
        //            // create the array of cells to add to the row:
        //            XElement[] cells = new XElement[_properties.Count];
        //            int c = 0;
        //            for (int _counter = 0; _counter < _properties.Count; _counter++)
        //            {
        //                if (_properties[_counter].GetValue(row) != null)
        //                {
        //                    switch (_properties[_counter].GetValue(row).GetType().ToString())
        //                    {
        //                        case "System.String":
        //                        case "System.Int16":
        //                        case "System.Int32":
        //                        case "System.Int64":
        //                        case "System.Byte":
        //                        case "System.Decimal":
        //                        case "System.Double":
        //                        case "System.DateTime":
        //                            cells[c++] =
        //                                new XElement("Cell",
        //                                    new XElement("Data", new XAttribute("ss:Type", "String"),
        //                                        new XText(_properties[_counter].GetValue(row).ToString())));
        //                            break;
        //                        default:
        //                            cells[c++] =
        //                                new XElement("Cell",
        //                                    new XElement("Data", new XAttribute("ss:Type", "String"),
        //                                        new XText(Resources.DefaultTypeText)));
        //                            break;
        //                    }
        //                }
        //                else
        //                    cells[c++] =
        //                        new XElement("Cell",
        //                            new XElement("Data", new XAttribute("ss:Type", "String"),
        //                                new XText(String.Empty)));
        //            }
        //            rows[r++] = new XElement("Row", cells);
        //        }
        //        return rows;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //    // return the array of rows.

        //}
    }

    /// <summary>
    /// Clase que define los tipos de hojas a utilizar para la exportación de colección de datos a PDF
    /// </summary>
    public class EnumExport
    {
        /// <summary>
        /// Lista de tipos de hoja a utilizar por el componente iTextSharp
        /// </summary>
        public enum TypeSheet
        {
            /// <summary>
            /// Tipo de hoja tamaño Carta
            /// </summary>
            Letter = 1,
            /// <summary>
            /// Tipo de hoja tamaño Carta con orientación horizontal
            /// </summary>
            LetterHorizontal = 2,
            /// <summary>
            /// Tipo de hoja tamaño A4
            /// </summary>
            A4 = 3,
            /// <summary>
            /// Tipo de hoja tamaño A4 con orientación horizontal
            /// </summary>
            A4Horizontal = 4,
            /// <summary>
            /// Tipo de hoja tamaño Oficio
            /// </summary>
            Legal = 5,
            /// <summary>
            /// Tipo de hoja tamaño oficio con orientación horizontal
            /// </summary>
            LegalHorizontal = 6
        }
    }
}
