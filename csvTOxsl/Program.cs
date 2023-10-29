namespace csvTOxlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            //Crea instancia para usar los metodos
            metodos metodo = new();

            //Asingacion de variables
            string ficheroCsv = string.Empty;
            string ficheroExcel = string.Empty;
            string plantillaExcel = string.Empty;
            int hoja = 0;//En la biblioteca NPOI el indice 0 es la primera hoja
            string celdaDestino = "A1";//Por defecto se pondra en la celda A1
            int fila = 1;
            int columna = 1;
            string textoLog = string.Empty;

            //Procesado de argumentos
            foreach (string arg in args)
            {
                if (arg == "-h")
                {
                    mensajeAyuda();
                    return;
                }
                string[] parts = arg.Split('=');
                if (parts.Length == 2)
                {
                    string key = parts[0].ToLower();
                    string value = parts[1];

                    switch (key)
                    {
                        case "entrada":
                            ficheroCsv = value;
                            break;

                        case "salida":
                            ficheroExcel = value;
                            break;

                        case "plantilla":
                            plantillaExcel = value;
                            break;

                        case "celda":
                            celdaDestino = value.ToUpper();
                            int[] columnaFila = metodo.convertirReferencia(celdaDestino);
                            fila = columnaFila[1];
                            columna = columnaFila[0];
                            break;

                        case "hoja":
                            hoja = Convert.ToInt32(value);
                            break;
                    }
                }
            }

            //Comprueba que se han pasado como parametro el fichero csv y el nombre del fichero de salida (la plantilla es opcional)
            if (string.IsNullOrEmpty(ficheroCsv) && string.IsNullOrEmpty(ficheroExcel))
            {
                textoLog += "Parametros incorrectos. No se ha informado del fichero CSV o el fichero de salida.\n";
                grabaResultado(textoLog);
                return;
            }

            if (hoja < 1)
            {
                textoLog += "El numero de hoja no puede ser menor de 1";
                grabaResultado(textoLog);
            }

            if (File.Exists(ficheroCsv))
            {
                try
                {
                    List<List<object>> datos = metodo.leerCSV(ficheroCsv); //Leer el archivo CSV
                    textoLog += metodo.exportaXLSX(datos, plantillaExcel, fila, columna, hoja, ficheroExcel); //Grabar el fichero Excel
                    //grabaResultado(textoLog);
                }
                catch (Exception ex)
                {
                    textoLog += "Error al procesar los ficheros: " + ex.Message + "\n";
                    grabaResultado(textoLog);
                }
            }
            else
            {
                textoLog += "Los archivos de entrada no estan en la carpeta.\n";
                grabaResultado(textoLog);
            }
        }

        private static void grabaResultado(string textoLog)
        {
            //Genera un fichero con el resultado
            string loggerFich = "resultado.txt";
            using (StreamWriter logger = new(loggerFich))
            {
                logger.WriteLine(textoLog);
            }
        }

        static void mensajeAyuda()
        {
            Console.Clear();
            Console.WriteLine("\nUso de la aplicacion.\n");
            Console.WriteLine("csvTOexcel [parametro1 parametro2 ... parametroN]\n");
            Console.WriteLine("Parametros:");
            Console.WriteLine("\tentrada=archivo.csv (obligatorio)\n\tsalida=archivo.xlsx (obligatorio)\n\tplantilla=plantilla.xlsx (opcional)\n\tcelda=A1 (defecto)\n\thoja=1 (defecto)");
        }
    }
}