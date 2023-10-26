using System.Data;
using Excel = OfficeOpenXml;
using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 3)
        {
            mensajeAyuda();
        }

        string ficheroCsv = string.Empty;
        string plantillaExcel = string.Empty;
        string celdaDestino = string.Empty;

        foreach (string arg in args)
        {
            string[] parts = arg.Split('=');
            if (parts.Length == 2)
            {
                string key = parts[0].ToLower();
                string value = parts[1];

                if (key == "origen")
                {
                    ficheroCsv = value;
                }
                else if (key == "plantilla")
                {
                    plantillaExcel = value;
                }
                else if (key == "celdadestino")
                {
                    celdaDestino = value.ToUpper();
                }
            }
        }

        if (string.IsNullOrEmpty(ficheroCsv) || string.IsNullOrEmpty(plantillaExcel) || string.IsNullOrEmpty(celdaDestino))
        {
            Console.WriteLine("Faltan parámetros o algunos son incorrectos.");
            mensajeAyuda();
            return;
        }

        if (File.Exists(ficheroCsv) && File.Exists(plantillaExcel))
        {
            try
            {
                DataTable tabla = ReadCsvToDataTable(ficheroCsv);
                WriteDataToExcel(tabla, plantillaExcel, celdaDestino);
                //Console.WriteLine("Datos insertados correctamente en Excel.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
        else
        {
            Console.WriteLine("El archivo CSV o la plantilla de Excel no existen en la ruta especificada.");
            mensajeAyuda();
        }


        static DataTable ReadCsvToDataTable(string filePath)
        {
            //Crea una tabla para insertar los datos que se extraigan del csv
            DataTable csvData = new DataTable();

            //Se crea una instancia TextFieldParser para tratar el archivo csv
            using (TextFieldParser parser = new TextFieldParser(filePath))
            {
                //Se le indica que los datos estan delimitados y separados por un punto y coma
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");
                parser.HasFieldsEnclosedInQuotes = true;

                // Lee los datos del csv
                if (!parser.EndOfData)
                {
                    string[] campos = parser.ReadFields();
                    for (int i = 0; i < campos.Length; i++)
                    {
                        csvData.Columns.Add(new DataColumn());
                    }

                    //string[] rowData = parser.ReadFields();
                    DataRow row = csvData.NewRow();
                    for (int i = 0; i < campos.Length; i++)
                    {
                        row[i] = campos[i];
                    }
                    csvData.Rows.Add(row);
                }

                while (!parser.EndOfData)
                {
                    string[] rowData = parser.ReadFields();
                    DataRow row = csvData.NewRow();
                    for (int i = 0; i < rowData.Length; i++)
                    {
                        row[i] = rowData[i];
                    }
                    csvData.Rows.Add(row);
                }
            }

            return csvData;
        }

        static void WriteDataToExcel(DataTable tabla, string plantillaExcel, string celdaDestino)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new Excel.ExcelPackage(new FileInfo(plantillaExcel)))
            {
                Excel.ExcelWorksheet hojaExcel = package.Workbook.Worksheets[0];

                int fila = int.Parse(celdaDestino.Substring(1));
                char columna = celdaDestino[0];

                int columnaIndex = columna - 'A'; // Calcula el índice de la columna en base a la letra

                for (int i = 0; i < tabla.Rows.Count; i++)
                {
                    for (int j = 0; j < tabla.Columns.Count; j++)
                    {
                        string contenidoCelda = tabla.Rows[i][j].ToString();
                        contenidoCelda = contenidoCelda.Replace(";", ",");
                        if (!string.IsNullOrEmpty(contenidoCelda) && contenidoCelda.StartsWith("="))
                        {
                            hojaExcel.Cells[fila + i, columnaIndex + j + 1].Formula = contenidoCelda; // +1 porque las columnas de Excel comienzan desde 1
                        }
                        else
                        {
                            hojaExcel.Cells[fila + i, columnaIndex + j + 1].Value = contenidoCelda;
                        }
                    }
                }
                //package.Workbook.Calculate();
                package.Save();
            }
        }

        static void mensajeAyuda()
        {
            Console.WriteLine("Parametros incorrectos. Se deben facilitar los siguientes parametros:");
            Console.WriteLine("csvTOxls origen=archivo.csv plantilla=plantilla.xls celdaDestino=A1");
        }
    }
}
