using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SimpleDataImport
{
    /// <summary>
    /// Class to easily import data from OpenXml, i.e Excel, documents 
    /// when used as a plain CSV database. That is the sole purpose of
    /// this lib.
    ///
    /// Example use:
    ///   var source = new OpenXmlImport<NamedEntry>(path);
    ///   List<NamedEntry> data = source.Import();
    ///
    /// NamedEntry is a class with properties named the same as the
    /// column names in the Excel file to be read. Data will
    /// be converted to appropriate data type on import.
    /// </summary>
    public class OpenXmlImport<T> where T: new()
    {
        private string fileName;
        private string sheetName;
        private bool hasColumnNames;
        private bool exitOnMissingColumn;

        public OpenXmlImport(string fileName, string sheetName = "Sheet1", bool hasColumnNames=true, bool exitOnMissingColumn=false)
        {
            this.fileName = fileName;
            this.sheetName = sheetName;
            this.hasColumnNames = hasColumnNames;
            this.exitOnMissingColumn = exitOnMissingColumn;
        }

        public List<T> Import()
        {
            var data = new List<T>();

            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(this.fileName, true))
            {
                WorkbookPart workbookPart = myDoc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == this.sheetName).FirstOrDefault();

                if (sheet == null)
                    throw new ArgumentException("Sheet name not found.");

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                Dictionary<char, string> map = new Dictionary<char, string>();
                bool readtitles = true;

                foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                {
                    Dictionary<string, object> newrow = new Dictionary<string, object>();

                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        if (cell == null)
                            continue;

                        string value = ParseCell(myDoc, cell);

                        if (readtitles)
                            if (this.hasColumnNames)
                                map.Add(cell.CellReference.Value[0], value);
                            else
                                map.Add(cell.CellReference.Value[0], cell.CellReference.Value[0].ToString());
                        else
                            newrow.Add(map[cell.CellReference.Value[0]], value);
                    }

                    if (readtitles)
                        readtitles = false;
                    else
                        data.Add(create(newrow));

                } // row
            }

            return data;
        }

        private T create(Dictionary<string, object> input)
        {
            var instance = new T();

            foreach (string key in input.Keys)
            {
                PropertyInfo propinfo = typeof(T).GetProperty(key);

                if (propinfo == null)
                    if (exitOnMissingColumn)
                        throw new Exception(string.Format("No {0} property on type {1}", key, typeof(T)));
                    else
                        continue;

                Type t = propinfo.PropertyType;
                var value = input[key];
                
                //var culture = new CultureInfo("en-US"); // Default formating seems to be en-US
                var culture = System.Globalization.CultureInfo.CurrentUICulture;
                
                if (t == typeof(string))
                    propinfo.SetValue(instance, value);
                if (t == typeof(Single))
                    propinfo.SetValue(instance, Convert.ToSingle(value, culture));
                if (t == typeof(double))
                    propinfo.SetValue(instance, Convert.ToDouble(value, culture));
                if (t == typeof(bool))
                    propinfo.SetValue(instance, (string)value == "1");
                if (t == typeof(int))
                    propinfo.SetValue(instance, Convert.ToInt32(value));
                if (t == typeof(DateTime))
                    propinfo.SetValue(instance, Convert.ToDateTime(value, culture));
            }

            return instance;
        }

        private string ParseCell(SpreadsheetDocument myDoc, Cell cell)
        {
            string value = cell.InnerText;

            if (cell.DataType != null)
            {
                if (cell.DataType.Value == CellValues.SharedString)
                {
                    SharedStringTable stringTable = myDoc.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    if (stringTable != null)
                        value = stringTable.ElementAt(int.Parse(value)).InnerText;
                }

                if (cell.DataType.Value == CellValues.Boolean)
                {
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        case "1":
                            value = "TRUE";
                            break;
                    }
                }
            }

            return value;
        }
    }
}
    
