using contifico.Models;
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
        private static readonly string apiKey = "FrguR1kDpFHaXHLQwplZ2CwTX3p8p9XHVTnukL98V5U";
        private static readonly string apiToken = "dce704ae-189e-4545-bea3-257d9249a594";
        private static readonly string endpointUrl = "https://api.contifico.com/sistema/api/v1/documento/";

        // Source and target folder paths (from configuration settings)
        private static readonly string folderPath = ConfigurationManager.AppSettings["SourcefolderPath"];
        private static readonly string folderBPath = ConfigurationManager.AppSettings["TargetfolderPath"];

        static async Task Main(string[] args)
        {
            // Ensure the source folder exists
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"Error: Source folder '{folderPath}' does not exist.");
                return;
            }

            // Retrieve all Excel files from the source folder
            string[] files = Directory.GetFiles(folderPath, "*.xlsx");

            // Group files based on numeric ID found in their names
            var groupedFiles = files.GroupBy(f => Regex.Match(Path.GetFileName(f), "\\d+").Value)
                          .ToDictionary(g => g.Key, g => g.ToList());

            if (groupedFiles.Count == 0)
            {
                Console.WriteLine("❌ No files matched the expected naming pattern.");
                return;
            }

            List<string> missingFilesLog = new List<string>();

            // Process each file pair
            foreach (var pair in groupedFiles)
            {
                Console.WriteLine($"Processing pair with ID: {pair.Key}");

                // Find 'detalle' and 'pedido' files
                string detalleFile = pair.Value.FirstOrDefault(f => f.Contains("fa_detalle_pedido"));
                string pedidoFile = pair.Value.FirstOrDefault(f => f.Contains("fa_Pedido"));

                // Log missing files
                if (detalleFile == null)
                    missingFilesLog.Add($"fa_detalle_pedido_{pair.Key}");

                if (pedidoFile == null)
                    missingFilesLog.Add($"fa_Pedido_{pair.Key}");

                // Skip processing if any required file is missing
                if (detalleFile == null || pedidoFile == null)
                    continue;

                // Read Excel data
                List<Detalle> detalles = ReadExcelData(detalleFile);
                List<Cliente> pedidos = ReadExcelDataPedido(pedidoFile);

                // Process data if successfully extracted
                if (detalles.Count > 0 && pedidos.Count > 0)
                {
                    await CreateDocumentAsync(detalles, pedidos, detalleFile, pedidoFile);
                }
                else
                {
                    Console.WriteLine($"Skipping pair {pair.Key}: Data extraction failed.");
                }
            }

            // 🔥 NEW: Display missing file report at the end
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

        // Reads 'detalle' file and extracts product details
        private static List<Detalle> ReadExcelData(string filePath)
        {
            var detalles = new List<Detalle>();
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return detalles;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int colCount = worksheet.Dimension?.Columns ?? 0;

                    if (rowCount == 0 || colCount == 0)
                    {
                        Console.WriteLine("Error: The worksheet is empty.");
                        return detalles;
                    }

                    // Read header names and map them to column indexes
                    var headers = new Dictionary<string, int>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string header = worksheet.Cells[1, col].Text.Trim().ToLower();
                        if (!string.IsNullOrEmpty(header))
                            headers[header] = col;
                    }

                    // Extract data from rows
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var detalle = new Detalle
                        {
                            producto_id = headers.ContainsKey("depe_codigo_producto") ? worksheet.Cells[row, headers["depe_codigo_producto"]].Text : "",
                            cantidad = headers.ContainsKey("depe_cantidad") && double.TryParse(worksheet.Cells[row, headers["depe_cantidad"]].Text, out double qty) ? qty : 0,
                            precio = headers.ContainsKey("depe_precio") && double.TryParse(worksheet.Cells[row, headers["depe_precio"]].Text, out double price) ? price : 0,
                            porcentaje_iva = headers.ContainsKey("depe_pago_iva") && int.TryParse(worksheet.Cells[row, headers["depe_pago_iva"]].Text, out int iva) ? iva : 0,
                            porcentaje_descuento = headers.ContainsKey("porcentaje_descuento") && double.TryParse(worksheet.Cells[row, headers["porcentaje_descuento"]].Text, out double descuento) ? descuento : 0,
                            base_cero = headers.ContainsKey("base_cero") && double.TryParse(worksheet.Cells[row, headers["base_cero"]].Text, out double base0) ? base0 : 0,
                            base_gravable = headers.ContainsKey("depe_precio") && double.TryParse(worksheet.Cells[row, headers["depe_precio"]].Text, out double baseGrav) ? baseGrav : 0,
                            base_no_gravable = headers.ContainsKey("base_no_gravable") && double.TryParse(worksheet.Cells[row, headers["base_no_gravable"]].Text, out double baseNoGrav) ? baseNoGrav : 0
                        };
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
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int colCount = worksheet.Dimension?.Columns ?? 0;

                    if (rowCount < 2 || colCount == 0)
                    {
                        Console.WriteLine("Error: The worksheet is empty or missing data.");
                        return clientes;
                    }

                    // Read header names and map them to column indexes
                    var headers = new Dictionary<string, int>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string header = worksheet.Cells[1, col].Text.Trim().ToLower();
                        if (!string.IsNullOrEmpty(header))
                            headers[header] = col;
                    }

                    // Extract data from rows
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var cliente = new Cliente
                        {
                            ruc = headers.ContainsKey("ruc") ? worksheet.Cells[row, headers["ruc"]].Text : "",
                            cedula = headers.ContainsKey("pedi_codigo_cliente") ? worksheet.Cells[row, headers["pedi_codigo_cliente"]].Text : "",
                            razon_social = headers.ContainsKey("pedi_nombre_cliente") ? worksheet.Cells[row, headers["pedi_nombre_cliente"]].Text : "",
                            telefonos = headers.ContainsKey("telefonos") ? worksheet.Cells[row, headers["telefonos"]].Text : "",
                            direccion = headers.ContainsKey("direccion") ? worksheet.Cells[row, headers["direccion"]].Text : "",
                            tipo = headers.ContainsKey("pedi_tipo") ? worksheet.Cells[row, headers["pedi_tipo"]].Text : "",
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
        private static async Task CreateDocumentAsync(List<Detalle> detalles, List<Cliente> pedidos,string detalleFile, string pedidoFile)
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

                var dummyData = new Documento
                {
                    pos = apiToken,
                    fecha_emision = "26/01/2016",
                    tipo_documento = "PRE",
                    estado = "P",
                    caja_id = "",
                    cliente = new Cliente
                    {
                        ruc = cliente.ruc,
                        cedula ="0" + cliente.cedula,
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
                    iva = detalles.Sum(d => d.base_gravable) * 0.12,
                    total = detalles.Sum(d => d.base_gravable) * 1.12,
                    adicional1 = string.Join("/", detalles.Select(d => d.producto_id)) + "/",
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
                    MoveFileToFolderB(pedidoFile);
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
