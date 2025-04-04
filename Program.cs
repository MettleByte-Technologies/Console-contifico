﻿using contifico.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace contifico
{
    internal class Program
    {
        // API credentials and endpoint URL
        private static readonly string apiKey = ConfigurationManager.AppSettings["apiKey"]; //"FrguR1kDpFHaXHLQwplZ2CwTX3p8p9XHVTnukL98V5U";;
        private static readonly string apiToken = ConfigurationManager.AppSettings["apiToken"]; //"dce704ae-189e-4545-bea3-257d9249a594";
        private static readonly string endpointUrl = ConfigurationManager.AppSettings["apiUrl"]; //"https://api.contifico.com/sistema/api/v1/documento/";

        // Source and target folder paths (from configuration settings)
        private static readonly string folderPath = ConfigurationManager.AppSettings["SourcefolderPath"];
        private static readonly string folderBPath = ConfigurationManager.AppSettings["TargetfolderPath"];

        static async Task Main(string[] args)
        {
            // Check if folderPath is valid
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: Source folder '{folderPath}' does not exist or is invalid.");
                return;
            }
            // Check if folderBPath is valid
            if (string.IsNullOrEmpty(folderBPath) || !Directory.Exists(folderBPath))
            {
                Console.WriteLine($"Error: Target folder '{folderBPath}' does not exist or is invalid.");
                return;
            }
            // Retrieve all Excel files from the source folder
            string[] files = Directory.GetFiles(folderPath, "*.xlsx");
            if (files.Length == 0)
            {
                Console.WriteLine(":x: No files found in the source folder.");
                return;
            }
            List<string> missingFilesLog = new List<string>();
            // Process each file
            foreach (var file in files)
            {
                Console.WriteLine($"Processing file: {file}");
                // Read Excel data
                string fetcha;
                List<Detalle> detalles = ReadExcelData(file, out fetcha);
                List<Cliente> pedidos = ReadExcelDataPedido(file);
                // Process data if successfully extracted
                if (detalles.Count > 0 && pedidos.Count > 0)
                {
                    await CreateDocumentAsync(detalles, pedidos, file, file, fetcha);
                }
                else
                {
                    Console.WriteLine($"Skipping file {file}: Data extraction failed.");
                }
            }
            // :fire: NEW: Display missing file report at the end
            if (missingFilesLog.Count > 0)
            {
                Console.WriteLine("\nMissing Files Report:");
                missingFilesLog.ForEach(Console.WriteLine);
            }
            else
            {
                Console.WriteLine("All required file pairs are present.");
            }
        }

        private static Dictionary<string, int> ExtractHeaders(ExcelWorksheet worksheet)
        {
            var headers = new Dictionary<string, int>();
            int colCount = worksheet.Dimension?.Columns ?? 0;

            for (int col = 1; col <= colCount; col++)
            {
                string header = worksheet.Cells[1, col].Text.Trim().ToLower();
                if (!string.IsNullOrEmpty(header))
                    headers[header] = col;
            }

            return headers;
        }

        // Reads 'detalle' file and extracts product details
        private static List<Detalle> ReadExcelData(string filePath, out string fecha)
        {
            var detalles = new List<Detalle>();
            fecha = "";
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return detalles;
            }

            try
            {
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage.License.SetNonCommercialPersonal("My Name");
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension?.Rows ?? 0;


                    if (rowCount == 0)
                    {
                        Console.WriteLine("Error: The worksheet is empty.");
                        return detalles;
                    }

                    // Read header names and map them to column indexes
                    var headers = ExtractHeaders(worksheet);  // Extract headers once
                    if (headers.ContainsKey("fecha_emision") && worksheet.Cells[2, headers["fecha_emision"]] != null)
                        fecha = worksheet.Cells[2, headers["fecha_emision"]].Text ?? "";
                    // Extract data from rows
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var detalle = new Detalle
                        {
                            producto_id = headers.ContainsKey("producto_id") ? worksheet.Cells[row, headers["producto_id"]].Text : "",
                            cantidad = headers.ContainsKey("cantidad") && double.TryParse(worksheet.Cells[row, headers["cantidad"]].Text, out double qty) ? qty : 0,
                            precio = headers.ContainsKey("precio") && double.TryParse(worksheet.Cells[row, headers["precio"]].Text, out double price) ? price : 0,
                            porcentaje_iva = headers.ContainsKey("porcentaje_iva") && int.TryParse(worksheet.Cells[row, headers["porcentaje_iva"]].Text, out int iva) ? iva : 0,
                            porcentaje_descuento = headers.ContainsKey("porcentaje_descuento") && double.TryParse(worksheet.Cells[row, headers["porcentaje_descuento"]].Text, out double descuento) ? descuento : 0,
                            base_cero = headers.ContainsKey("base_cero") && double.TryParse(worksheet.Cells[row, headers["base_cero"]].Text, out double base0) ? base0 : 0,
                            base_gravable = 0,
                            base_no_gravable = headers.ContainsKey("base_no_gravable") && double.TryParse(worksheet.Cells[row, headers["base_no_gravable"]].Text, out double baseNoGrav) ? baseNoGrav : 0
                        };
                        // ✅ Set base_gravable as precio * cantidad
                        if (detalle.porcentaje_iva == 0)
                        {
                            detalle.base_cero = detalle.precio * detalle.cantidad;  // Assign base_cero when tax is 0
                            detalle.base_gravable = 0;
                        }
                        else
                        {
                            double descuento_aplicado = (detalle.porcentaje_descuento / 100) * detalle.precio * detalle.cantidad;
                            detalle.base_gravable = (detalle.precio * detalle.cantidad) - descuento_aplicado;
                            detalle.base_cero = 0;
                        }

                        detalles.Add(detalle);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
            }
            return detalles;
        }

        // Reads 'pedido' file and extracts client details
        private static List<Cliente> ReadExcelDataPedido(string filePath)
        {
            var clientes = new List<Cliente>();
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return clientes;
            }

            try
            {
                // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage.License.SetNonCommercialPersonal("My Name");
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension?.Rows ?? 0;



                    if (rowCount < 2)
                    {
                        Console.WriteLine("Error: The worksheet is empty or missing data.");
                        return clientes;
                    }

                    // Read header names and map them to column indexes
                    var headers = ExtractHeaders(worksheet);  // Reuse the same header extraction function


                    // Extract data from rows
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var cliente = new Cliente
                        {
                            ruc = headers.ContainsKey("ruc") ? worksheet.Cells[row, headers["ruc"]].Text : "",
                            cedula = headers.ContainsKey("cedula") ? worksheet.Cells[row, headers["cedula"]].Text : "",
                            razon_social = headers.ContainsKey("razon_social") ? worksheet.Cells[row, headers["razon_social"]].Text : "",
                            telefonos = headers.ContainsKey("telefonos") ? worksheet.Cells[row, headers["telefonos"]].Text : "",
                            direccion = headers.ContainsKey("direccion") ? worksheet.Cells[row, headers["direccion"]].Text : "",
                            tipo = headers.ContainsKey("tipo") ? worksheet.Cells[row, headers["tipo"]].Text : "",
                            email = headers.ContainsKey("email") ? worksheet.Cells[row, headers["email"]].Text : "",
                            es_extranjero = headers.ContainsKey("es_extranjero") && worksheet.Cells[row, headers["es_extranjero"]].Text.ToLower() == "true"
                        };
                        clientes.Add(cliente);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Excel file: {ex.Message}");
            }
            return clientes;
        }

        // Creates a document using API call
        private static async Task CreateDocumentAsync(List<Detalle> detalles, List<Cliente> pedidos, string detalleFile, string pedidoFile, string fecha)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Api-Key", apiKey);
                client.DefaultRequestHeaders.Add("Authorization", apiKey);

                var cliente = pedidos.FirstOrDefault(); // Take first Cliente object

                if (cliente == null)
                {
                    Console.WriteLine("❌ No client data available. API call aborted.");
                    return;
                }
                //string formattedCedula = cliente.cedula.Length == 9 ? "0" + cliente.cedula : cliente.cedula;
                var dummyData = new Documento
                {
                    pos = apiToken,
                    fecha_emision = fecha.Replace("-", "/"),
                    tipo_documento = "PRE",
                    estado = "P",
                    caja_id = "",
                    cliente = new Cliente
                    {
                        ruc = cliente.ruc,
                        cedula = cliente.cedula,
                        razon_social = cliente.razon_social,
                        telefonos = cliente.telefonos,
                        direccion = cliente.direccion,
                        tipo = cliente.tipo,
                        email = cliente.email,
                        es_extranjero = cliente.es_extranjero
                    },
                    vendedor = "",
                    descripcion = "DETALLE PREFACTURA",
                    subtotal_0 = detalles.Sum(d => d.base_cero),
                    subtotal_12 = detalles.Sum(d => d.base_gravable),
                    iva = detalles.Sum(d => d.base_gravable * (d.porcentaje_iva / 100.0)),
                    total = detalles.Sum(d => d.base_cero + d.base_gravable + (d.base_gravable * (d.porcentaje_iva / 100.0))),
                    adicional1 = "",
                    detalles = detalles.ToArray()
                };
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(dummyData, Newtonsoft.Json.Formatting.Indented);
                Console.WriteLine($"Request Payload: {json}"); // Debugging


                StringContent content = new StringContent(json, Encoding.UTF8, "application/json");


                HttpResponseMessage response = await client.PostAsync(endpointUrl, content);
                string responseContent = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"API Response: {responseContent}");

                if (response.IsSuccessStatusCode)
                {
                    MoveFileToFolderB(detalleFile);

                }
                else
                {
                    Console.WriteLine("❌ API call failed. Files will not be moved.");
                }
            }
        }

        // Moves processed files to the target folder
        private static void MoveFileToFolderB(string filePath)
        {
            string folderBPath = ConfigurationManager.AppSettings["TargetfolderPath"];

            if (!Directory.Exists(folderBPath))
            {
                Directory.CreateDirectory(folderBPath);
            }

            string newFileName = Path.Combine(folderBPath, Path.GetFileNameWithoutExtension(filePath) + "_old.xlsx");

            try
            {
                File.Move(filePath, newFileName);
                Console.WriteLine($"✅ Moved file to: {newFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error moving file: {ex.Message}");
            }
        }
    }
}
