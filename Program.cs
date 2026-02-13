using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Text;
using System.IO.Ports;
using System.Management;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace RawPrint
{
    public class PrinterService : ServiceBase
    {
        private HttpListener _listener;
        private CancellationTokenSource _cts;

        public PrinterService()
        {
            ServiceName = "RawPrintService";
        }

        protected override void OnStart(string[] args)
        {
            _cts = new CancellationTokenSource();
            Task.Run(() => RunServer(_cts.Token));
        }

        protected override void OnStop()
        {
            _cts.Cancel();
            _listener?.Stop();
        }

        private void RunServer(CancellationToken token)
        {
            try
            {
                _listener = new HttpListener();
                _listener.Prefixes.Add("http://+:8100/");
                _listener.Start();

                while (!token.IsCancellationRequested)
                {
                    var ctx = _listener.GetContext();
                    Task.Run(() => HandleRequest(ctx));
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry($"Error: {ex.Message}", EventLogEntryType.Error);
            }
        }

        private void HandleRequest(HttpListenerContext ctx)
        {
            try
            {
                // Normalize path - fix trailing slash and case issues
                string path = ctx.Request.Url.AbsolutePath.TrimEnd('/').ToLower();
                string method = ctx.Request.HttpMethod.ToUpper();

                // Debug log
                Console.WriteLine($"[REQUEST] {method} {path}");

                // Handle preflight CORS
                if (method == "OPTIONS")
                {
                    ctx.Response.AppendHeader("Access-Control-Allow-Origin", "*");
                    ctx.Response.AppendHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                    ctx.Response.AppendHeader("Access-Control-Allow-Headers", "Content-Type, Accept");
                    ctx.Response.StatusCode = 200;
                    ctx.Response.OutputStream.Close();
                    return;
                }

                if (path == "" || path == "/")
                {
                    var computer = Environment.MachineName;
                    var ip = Dns.GetHostAddresses(computer)
                                .FirstOrDefault(ip => ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);

                    var printers = PrinterSettings.InstalledPrinters.Cast<string>().ToList();

                    var response = new
                    {
                        Computer = computer,
                        IP = ip?.ToString(),
                        Printers = printers
                    };

                    string json = JsonSerializer.Serialize(response, new JsonSerializerOptions { WriteIndented = true });
                    WriteResponse(ctx, json, "application/json");
                }
                else if (path == "/print" && method == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();
                        var doc = JsonSerializer.Deserialize<Dictionary<string, string>>(body);

                        if (!doc.ContainsKey("printer") || !doc.ContainsKey("data"))
                        {
                            WriteResponse(ctx, "{\"error\":\"Missing printer or data\"}", "application/json");
                            return;
                        }

                        var printerName = doc["printer"];
                        var data = doc["data"];
                        var encoding = doc.ContainsKey("encoding") ? doc["encoding"] : "ascii";

                        try
                        {
                            byte[] bytes;
                            if (encoding.ToLower() == "base64")
                                bytes = Convert.FromBase64String(data);
                            else
                                bytes = Encoding.ASCII.GetBytes(data);

                            PrintRawBytes(printerName, bytes);
                            WriteResponse(ctx, "{\"status\":\"sent to printer\"}", "application/json");
                        }
                        catch (Exception ex)
                        {
                            WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
                        }
                    }
                }
                else if (path == "/print-image" && method == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();

                        try
                        {
                            using (JsonDocument jsonDoc = JsonDocument.Parse(body))
                            {
                                var root = jsonDoc.RootElement;

                                if (!root.TryGetProperty("printer", out var printerElement) ||
                                    !root.TryGetProperty("image", out var imageElement))
                                {
                                    WriteResponse(ctx, "{\"error\":\"Missing printer or image data\"}", "application/json");
                                    return;
                                }

                                var printerName = printerElement.GetString();
                                var imageBase64 = imageElement.GetString();

                                var fitToPage = root.TryGetProperty("fitToPage", out var fitElement)
                                    ? fitElement.GetBoolean() : true;

                                float widthInches = root.TryGetProperty("widthInches", out var wElement)
                                    ? (float)wElement.GetDouble() : 0;
                                float heightInches = root.TryGetProperty("heightInches", out var hElement)
                                    ? (float)hElement.GetDouble() : 0;

                                byte[] imageBytes = Convert.FromBase64String(imageBase64);
                                using (var ms = new MemoryStream(imageBytes))
                                using (var img = Image.FromStream(ms))
                                {
                                    PrintImage(printerName, img, fitToPage, widthInches, heightInches);
                                }

                                WriteResponse(ctx, "{\"status\":\"image sent to printer\"}", "application/json");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[ERROR] print-image: {ex.Message}");
                            WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
                        }
                    }
                }
                else if (path == "/print-pdf" && method == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();

                        try
                        {
                            using (JsonDocument jsonDoc = JsonDocument.Parse(body))
                            {
                                var root = jsonDoc.RootElement;

                                if (!root.TryGetProperty("printer", out var printerElement) ||
                                    !root.TryGetProperty("pdf", out var pdfElement))
                                {
                                    WriteResponse(ctx, "{\"error\":\"Missing printer or pdf data\"}", "application/json");
                                    return;
                                }

                                var printerName = printerElement.GetString();
                                var pdfPath = pdfElement.GetString();

                                if (!File.Exists(pdfPath))
                                {
                                    WriteResponse(ctx, "{\"error\":\"PDF file not found\"}", "application/json");
                                    return;
                                }

                                PrintPdf(printerName, pdfPath);

                                WriteResponse(ctx, "{\"status\":\"pdf sent to printer\"}", "application/json");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[ERROR] print-pdf: {ex.Message}");
                            WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
                        }
                    }
                }
                // else if (path == "/print-html" && method == "POST")
                // {
                //     using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                //     {
                //         var body = reader.ReadToEnd();

                //         try
                //         {
                //             using (JsonDocument jsonDoc = JsonDocument.Parse(body))
                //             {
                //                 var root = jsonDoc.RootElement;

                //                 if (!root.TryGetProperty("printer", out var printerElement) ||
                //                     !root.TryGetProperty("html", out var htmlElement))
                //                 {
                //                     WriteResponse(ctx, "{\"error\":\"Missing printer or html\"}", "application/json");
                //                     return;
                //                 }

                //                 var printerName = printerElement.GetString();
                //                 var htmlContent = htmlElement.GetString();

                //                 PrintHtml(printerName, htmlContent);

                //                 WriteResponse(ctx, "{\"status\":\"html sent to printer\"}", "application/json");
                //             }
                //         }
                //         catch (Exception ex)
                //         {
                //             WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
                //         }
                //     }
                // }

                else if (path == "/print-html" && method == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream,
                        ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();

                        try
                        {
                            using (JsonDocument jsonDoc = JsonDocument.Parse(body))
                            {
                                var root = jsonDoc.RootElement;

                                if (!root.TryGetProperty("printer", out var printerEl) ||
                                    !root.TryGetProperty("html", out var htmlEl))
                                {
                                    WriteResponse(ctx,
                                        "{\"error\":\"Missing printer or html\"}",
                                        "application/json");
                                    return;
                                }

                                var printerName = printerEl.GetString();
                                var html = htmlEl.GetString();

                                float widthInches = root.TryGetProperty("widthInches", out var wEl)
                                    ? (float)wEl.GetDouble() : 8.5f;
                                float heightInches = root.TryGetProperty("heightInches", out var hEl)
                                    ? (float)hEl.GetDouble() : 5.5f;

                                // DPI settings for dot matrix
                                // dpiX: horizontal DPI (default 240)
                                // dpiY: vertical DPI (default 72)
                                int dpiX = root.TryGetProperty("dpiX", out var dxEl)
                                    ? dxEl.GetInt32() : 240;
                                int dpiY = root.TryGetProperty("dpiY", out var dyEl)
                                    ? dyEl.GetInt32() : 72;

                                PrintHtml(printerName, html, widthInches, heightInches, dpiX, dpiY);

                                WriteResponse(ctx,
                                    "{\"status\":\"html sent to printer\"}",
                                    "application/json");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[ERROR] print-html: {ex.Message}");
                            WriteResponse(ctx,
                                $"{{\"error\":\"{ex.Message}\"}}",
                                "application/json");
                        }
                    }
                }
                else if (path == "/status" && method == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();

                        try
                        {
                            using (JsonDocument jsonDoc = JsonDocument.Parse(body))
                            {
                                var root = jsonDoc.RootElement;

                                if (!root.TryGetProperty("printer", out var printerElement))
                                {
                                    WriteResponse(ctx, "{\"error\":\"Missing printer name\"}", "application/json");
                                    return;
                                }

                                var printerName = printerElement.GetString();
                                string status = GetPrinterStatus(printerName);

                                WriteResponse(ctx, $"{{\"status\":\"{status}\"}}", "application/json");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
                        }
                    }
                }
                else if (path == "/test" && method == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();

                        try
                        {
                            using (JsonDocument jsonDoc = JsonDocument.Parse(body))
                            {
                                var root = jsonDoc.RootElement;

                                if (!root.TryGetProperty("printer", out var printerElement))
                                {
                                    WriteResponse(ctx, "{\"error\":\"Missing printer name\"}", "application/json");
                                    return;
                                }

                                var printerName = printerElement.GetString();
                                string result = PrintTestPage(printerName);

                                var json = JsonSerializer.Serialize(new { status = "ok", message = result });
                                WriteResponse(ctx, json, "application/json");
                            }
                        }
                        catch (Exception ex)
                        {
                            var json = JsonSerializer.Serialize(new { error = ex.Message });
                            WriteResponse(ctx, json, "application/json");
                        }
                    }
                }
                else
                {
                    // Show exactly what path was not found
                    Console.WriteLine($"[NOT FOUND] {method} {path}");
                    ctx.Response.StatusCode = 404;
                    WriteResponse(ctx, $"{{\"error\":\"Not Found\", \"path\":\"{path}\", \"method\":\"{method}\"}}", "application/json");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] {ex.Message}");
                WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
            }
            finally
            {
                ctx.Response.OutputStream.Close();
            }
        }        // ---------- IMAGE PRINTING ----------

        private void PrintImage(string printerName, Image image, bool fitToPage, float widthInches = 0, float heightInches = 0)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = printerName;

            // Remove all margins
            pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            // If custom size provided, set paper size
            if (widthInches > 0 && heightInches > 0)
            {
                // Convert inches to hundredths of an inch (printer units)
                int paperWidth = (int)(widthInches * 100);
                int paperHeight = (int)(heightInches * 100);

                pd.DefaultPageSettings.PaperSize = new PaperSize("Custom", paperWidth, paperHeight);
            }

            pd.PrintPage += (sender, e) =>
            {
                // Use inches instead of pixels
                e.Graphics.PageUnit = GraphicsUnit.Inch;

                if (fitToPage)
                {
                    RectangleF area = new RectangleF(
                        0,
                        0,
                        e.PageBounds.Width / 100f,
                        e.PageBounds.Height / 100f
                    );

                    e.Graphics.DrawImage(image, area);
                }
                else
                {
                    // Draw EXACT physical size in inches
                    RectangleF exactSize = new RectangleF(
                        0,
                        0,
                        widthInches,
                        heightInches
                    );

                    e.Graphics.DrawImage(image, exactSize);
                }
            };


            pd.Print();
        }

        // ---------- PDF PRINTING ----------
        // private void PrintPdf(string printerName, byte[] pdfBytes)
        // {
        //     // Save PDF to temp file
        //     string tempPdfPath = Path.Combine(Path.GetTempPath(), $"print_{Guid.NewGuid()}.pdf");
        //     File.WriteAllBytes(tempPdfPath, pdfBytes);

        //     try
        //     {
        //         // Use Adobe Reader or default PDF viewer to print
        //         ProcessStartInfo psi = new ProcessStartInfo
        //         {
        //             FileName = tempPdfPath,
        //             Verb = "print",
        //             Arguments = $"/p /h \"{tempPdfPath}\"",
        //             CreateNoWindow = true,
        //             WindowStyle = ProcessWindowStyle.Hidden,
        //             UseShellExecute = true
        //         };

        //         // Alternative: Use SumatraPDF if installed (more reliable)
        //         string sumatraPath = @"C:\Program Files\SumatraPDF\SumatraPDF.exe";
        //         if (File.Exists(sumatraPath))
        //         {
        //             psi.FileName = sumatraPath;
        //             psi.Arguments = $"-print-to \"{printerName}\" \"{tempPdfPath}\"";
        //             psi.UseShellExecute = false;
        //         }

        //         using (Process p = Process.Start(psi))
        //         {
        //             if (p != null && !p.HasExited)
        //             {
        //                 p.WaitForExit(10000); // Wait up to 10 seconds
        //             }
        //         }

        //         // Wait a bit before deleting temp file
        //         Thread.Sleep(2000);
        //     }
        //     finally
        //     {
        //         // Clean up temp file
        //         try
        //         {
        //             if (File.Exists(tempPdfPath))
        //                 File.Delete(tempPdfPath);
        //         }
        //         catch { }
        //     }
        // }

        // private void PrintPdf(string printerName, byte[] pdfBytes)
        // {
        //     string tempPdfPath = Path.Combine(
        //         Path.GetTempPath(),
        //         $"print_{Guid.NewGuid()}.pdf"
        //     );

        //     File.WriteAllBytes(tempPdfPath, pdfBytes);

        //     try
        //     {
        //         string sumatraPath = @"C:\SumatraPDF\SumatraPDF.exe";

        //         if (!File.Exists(sumatraPath))
        //             throw new Exception("SumatraPDF not found.");

        //         ProcessStartInfo psi = new ProcessStartInfo
        //         {
        //             FileName = sumatraPath,
        //             Arguments = $"-print-to \"{printerName}\" -silent \"{tempPdfPath}\"",
        //             UseShellExecute = false,
        //             CreateNoWindow = true,
        //             WindowStyle = ProcessWindowStyle.Hidden
        //         };

        //         using (Process p = Process.Start(psi))
        //         {
        //             p?.WaitForExit(15000);
        //         }

        //         Thread.Sleep(1000);
        //     }
        //     finally
        //     {
        //         try
        //         {
        //             if (File.Exists(tempPdfPath))
        //                 File.Delete(tempPdfPath);
        //         }
        //         catch { }
        //     }
        // }

        private void PrintPdf(string printerName, string pdfPath)
        {
            string sumatraPath = @"C:\SumatraPDF\SumatraPDF.exe";

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = sumatraPath,
                Arguments = $"-print-to \"{printerName}\" \"{pdfPath}\"",
                UseShellExecute = false,
                CreateNoWindow = false
            };

            using (Process p = Process.Start(psi))
            {
                p?.WaitForExit(15000);
            }
        }

        // private void PrintHtml(string printerName, string html)
        // {
        //     var thread = new Thread(() =>
        //     {
        //         var form = new Form();
        //         form.WindowState = FormWindowState.Minimized;
        //         form.ShowInTaskbar = false;

        //         var webView = new WebView2();
        //         webView.Dock = DockStyle.Fill;
        //         form.Controls.Add(webView);

        //         form.Load += async (s, e) =>
        //         {
        //             await webView.EnsureCoreWebView2Async();

        //             webView.NavigateToString(html);

        //             webView.CoreWebView2.NavigationCompleted += async (sender, args) =>
        //             {
        //                 var settings = webView.CoreWebView2.Environment.CreatePrintSettings();
        //                 settings.ShouldPrintBackgrounds = true;
        //                 settings.PrinterName = printerName;

        //                 await webView.CoreWebView2.PrintAsync(settings);

        //                 form.Close();
        //             };
        //         };

        //         Application.Run(form);
        //     });

        //     thread.SetApartmentState(ApartmentState.STA);
        //     thread.Start();
        //     thread.Join();
        // }


        // Alternative PDF printing using GhostScript (requires GhostScript installed)

        // private void PrintHtml(string printerName, string html, float widthInches = 8.5f, float heightInches = 5.5f)
        // {
        //     string tempPdfPath = Path.Combine(Path.GetTempPath(),
        //         $"print_{Guid.NewGuid()}.pdf");

        //     ManualResetEventSlim done = new ManualResetEventSlim(false);
        //     Exception error = null;

        //     var thread = new Thread(() =>
        //     {
        //         Form form = null;
        //         WebView2 webView = null;

        //         try
        //         {
        //             form = new Form();
        //             form.WindowState = FormWindowState.Minimized;
        //             form.ShowInTaskbar = false;

        //             webView = new WebView2();
        //             webView.Dock = DockStyle.Fill;
        //             form.Controls.Add(webView);

        //             form.Load += async (s, e) =>
        //             {
        //                 try
        //                 {
        //                     await webView.EnsureCoreWebView2Async();
        //                     webView.NavigateToString(html);

        //                     webView.CoreWebView2.NavigationCompleted += async (sender, args) =>
        //                     {
        //                         try
        //                         {
        //                             // Wait for full render
        //                             await Task.Delay(800);

        //                             // ✅ Create VECTOR PDF - same as Chrome print!
        //                             var printSettings = webView.CoreWebView2
        //                                 .Environment.CreatePrintSettings();

        //                             printSettings.PageWidth = widthInches;
        //                             printSettings.PageHeight = heightInches;
        //                             printSettings.MarginTop = 0;
        //                             printSettings.MarginBottom = 0;
        //                             printSettings.MarginLeft = 0;
        //                             printSettings.MarginRight = 0;
        //                             printSettings.ScaleFactor = 1.0;
        //                             printSettings.ShouldPrintBackgrounds = false;
        //                             printSettings.ShouldPrintHeaderAndFooter = false;

        //                             // This creates VECTOR PDF - text stays as text!
        //                             await webView.CoreWebView2.PrintToPdfAsync(
        //                                 tempPdfPath, printSettings);

        //                             Console.WriteLine($"[INFO] Vector PDF created!");
        //                         }
        //                         catch (Exception ex)
        //                         {
        //                             error = ex;
        //                         }
        //                         finally
        //                         {
        //                             done.Set();
        //                             form?.Invoke(new Action(() => form?.Close()));
        //                         }
        //                     };
        //                 }
        //                 catch (Exception ex)
        //                 {
        //                     error = ex;
        //                     done.Set();
        //                     form?.Close();
        //                 }
        //             };

        //             Application.Run(form);
        //         }
        //         catch (Exception ex)
        //         {
        //             error = ex;
        //             done.Set();
        //         }
        //     });

        //     thread.SetApartmentState(ApartmentState.STA);
        //     thread.Start();
        //     done.Wait(TimeSpan.FromSeconds(30));
        //     thread.Join(TimeSpan.FromSeconds(5));

        //     if (error != null) throw error;
        //     if (!File.Exists(tempPdfPath))
        //         throw new Exception("PDF was not created");

        //     // Print the vector PDF
        //     try
        //     {
        //         PrintVectorPdf(printerName, tempPdfPath);
        //     }
        //     finally
        //     {
        //         Thread.Sleep(3000);
        //         try { File.Delete(tempPdfPath); } catch { }
        //     }
        // }

        private void PrintHtml(string printerName, string html,
    float widthInches = 8.5f, float heightInches = 5.5f,
    int dpiX = 240, int dpiY = 72)
        {
            string tempPdfPath = Path.Combine(Path.GetTempPath(),
                $"print_{Guid.NewGuid()}.pdf");

            ManualResetEventSlim done = new ManualResetEventSlim(false);
            Exception error = null;

            var thread = new Thread(() =>
            {
                Form form = null;
                WebView2 webView = null;

                try
                {
                    form = new Form();
                    form.WindowState = FormWindowState.Minimized;
                    form.ShowInTaskbar = false;

                    webView = new WebView2();
                    webView.Dock = DockStyle.Fill;
                    form.Controls.Add(webView);

                    form.Load += async (s, e) =>
                    {
                        try
                        {
                            await webView.EnsureCoreWebView2Async();
                            webView.NavigateToString(html);

                            webView.CoreWebView2.NavigationCompleted += async (sender, args) =>
                            {
                                try
                                {
                                    await Task.Delay(800);

                                    var printSettings = webView.CoreWebView2
                                        .Environment.CreatePrintSettings();

                                    printSettings.PageWidth = widthInches;
                                    printSettings.PageHeight = heightInches;
                                    printSettings.MarginTop = 0;
                                    printSettings.MarginBottom = 0;
                                    printSettings.MarginLeft = 0;
                                    printSettings.MarginRight = 0;
                                    printSettings.ScaleFactor = 1.0;
                                    printSettings.ShouldPrintBackgrounds = true;
                                    printSettings.ShouldPrintHeaderAndFooter = false;

                                    // Create vector PDF
                                    await webView.CoreWebView2.PrintToPdfAsync(
                                        tempPdfPath, printSettings);

                                    Console.WriteLine($"[INFO] Vector PDF created: {tempPdfPath}");
                                }
                                catch (Exception ex)
                                {
                                    error = ex;
                                    Console.WriteLine($"[ERROR] PDF creation failed: {ex.Message}");
                                }
                                finally
                                {
                                    done.Set();
                                    form?.Invoke(new Action(() => form?.Close()));
                                }
                            };
                        }
                        catch (Exception ex)
                        {
                            error = ex;
                            done.Set();
                            form?.Close();
                        }
                    };

                    Application.Run(form);
                }
                catch (Exception ex)
                {
                    error = ex;
                    done.Set();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            done.Wait(TimeSpan.FromSeconds(30));
            thread.Join(TimeSpan.FromSeconds(5));

            if (error != null) throw error;
            if (!File.Exists(tempPdfPath))
                throw new Exception("PDF was not created");

            try
            {
                PrintVectorPdf(printerName, tempPdfPath, dpiX, dpiY);
            }
            finally
            {
                Thread.Sleep(3000);
                try { File.Delete(tempPdfPath); } catch { }
            }
        }

        private void PrintVectorPdf(string printerName, string pdfPath,
            int dpiX = 240, int dpiY = 72)
        {
            string gsPath = FindGhostScript();

            if (gsPath == null)
                throw new Exception(
                    "GhostScript not found at " +
                    "C:\\Program Files\\gs\\gs10.06.0\\bin\\gswin64c.exe");

            Console.WriteLine($"[INFO] Printing via GhostScript ESC/P...");
            Console.WriteLine($"[INFO] Printer : {printerName}");
            Console.WriteLine($"[INFO] DPI     : {dpiX}x{dpiY}");
            Console.WriteLine($"[INFO] PDF     : {pdfPath}");

            var psi = new ProcessStartInfo
            {
                FileName = gsPath,
                Arguments = $"-q -dNOPAUSE -dSAFER -dBATCH " +
                           $"-sDEVICE=escp " +
                           $"-r{dpiX}x{dpiY} " +
                           $"-sOutputFile=\"%printer%{printerName}\" " +
                           $"\"{pdfPath}\"",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var p = Process.Start(psi))
            {
                string output = p.StandardOutput.ReadToEnd();
                string errors = p.StandardError.ReadToEnd();

                p?.WaitForExit(30000);

                if (!string.IsNullOrEmpty(output))
                    Console.WriteLine($"[GS OUTPUT] {output}");
                if (!string.IsNullOrEmpty(errors))
                    Console.WriteLine($"[GS ERROR] {errors}");

                Console.WriteLine($"[INFO] GhostScript exit code: {p.ExitCode}");
            }

            Console.WriteLine($"[INFO] Done printing to {printerName}");
        }
        private void PrintVectorPdf(string printerName, string pdfPath)
        {
            string gsPath = FindGhostScript();

            if (gsPath == null)
                throw new Exception("GhostScript not found at C:\\Program Files\\gs\\gs10.06.0\\bin\\gswin64c.exe");

            Console.WriteLine($"[INFO] Printing via GhostScript ESC/P to {printerName}...");

            var psi = new ProcessStartInfo
            {
                FileName = gsPath,
                Arguments = $"-q -dNOPAUSE -dSAFER -dBATCH " +
                           $"-sDEVICE=escp " +          // ESC/P device for dot matrix
                           $"-r240x144 " +               // 240x72 DPI (dot matrix standard)
                           $"-sOutputFile=\"%printer%{printerName}\" " +
                           $"\"{pdfPath}\"",
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var p = Process.Start(psi))
            {
                // Capture any errors
                string output = p.StandardOutput.ReadToEnd();
                string errors = p.StandardError.ReadToEnd();

                p?.WaitForExit(30000);

                if (!string.IsNullOrEmpty(output))
                    Console.WriteLine($"[GS OUTPUT] {output}");
                if (!string.IsNullOrEmpty(errors))
                    Console.WriteLine($"[GS ERROR] {errors}");

                Console.WriteLine($"[INFO] GhostScript exit code: {p.ExitCode}");
            }

            Console.WriteLine($"[INFO] Printed via GhostScript ESC/P to {printerName}!");
        }
        // private void PrintVectorPdf(string printerName, string pdfPath)
        // {
        //     // Try SumatraPDF first (best for dot matrix)
        //     string[] sumatraPaths = { @"C:\SumatraPDF\SumatraPDF.exe" };

        //     foreach (var path in sumatraPaths)
        //     {
        //         if (File.Exists(path))
        //         {
        //             Console.WriteLine("[INFO] Printing via SumatraPDF (vector)...");
        //             var psi = new ProcessStartInfo
        //             {
        //                 FileName = path,
        //                 // -print-settings "fit" ensures correct paper size
        //                 Arguments = $"-print-to \"{printerName}\" -print-settings \"fit\" \"{pdfPath}\"",
        //                 CreateNoWindow = true,
        //                 WindowStyle = ProcessWindowStyle.Hidden,
        //                 UseShellExecute = false
        //             };
        //             using (var p = Process.Start(psi))
        //             {
        //                 p?.WaitForExit(15000);
        //             }
        //             Console.WriteLine("[INFO] Printed via SumatraPDF!");
        //             return;
        //         }
        //     }

        //     // Try GhostScript
        //     string gsPath = FindGhostScript();
        //     if (gsPath != null)
        //     {
        //         Console.WriteLine("[INFO] Printing via GhostScript (vector)...");
        //         var psi = new ProcessStartInfo
        //         {
        //             FileName = gsPath,
        //             Arguments = $"-dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q " +
        //                        $"-dNumCopies=1 -sDEVICE=mswinpr2 " +
        //                        $"-sOutputFile=\"%printer%{printerName}\" \"{pdfPath}\"",
        //             CreateNoWindow = true,
        //             WindowStyle = ProcessWindowStyle.Hidden,
        //             UseShellExecute = false
        //         };
        //         using (var p = Process.Start(psi))
        //         {
        //             p?.WaitForExit(30000);
        //         }
        //         Console.WriteLine("[INFO] Printed via GhostScript!");
        //         return;
        //     }

        //     throw new Exception(
        //         "No PDF viewer found. Install SumatraPDF: " +
        //         "https://www.sumatrapdfreader.org/");
        // }

        private string FindGhostScript()
        {
            string[] paths = new[]
            {
                // Your installed version - 64-bit console (recommended - no window popup)
                @"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe",
                
                // Your installed version - 64-bit with window (fallback)
                @"C:\Program Files\gs\gs10.06.0\bin\gswin64.exe",
            };

            foreach (var path in paths)
            {
                if (File.Exists(path))
                {
                    Console.WriteLine($"[INFO] GhostScript found: {path}");
                    return path;
                }
            }

            Console.WriteLine("[ERROR] GhostScript not found!");
            return null;
        }

        private void PrintPdfWithGhostScript(string printerName, byte[] pdfBytes)
        {
            string tempPdfPath = Path.Combine(Path.GetTempPath(), $"print_{Guid.NewGuid()}.pdf");
            File.WriteAllBytes(tempPdfPath, pdfBytes);

            try
            {
                string gsPath = @"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe";

                if (!File.Exists(gsPath))
                {
                    throw new Exception("GhostScript not found. Please install GhostScript.");
                }

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = gsPath,
                    Arguments = $"-dPrinted -dBATCH -dNOPAUSE -dNOSAFER -q -dNumCopies=1 -sDEVICE=mswinpr2 -sOutputFile=\"%printer%{printerName}\" \"{tempPdfPath}\"",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = false
                };

                using (Process p = Process.Start(psi))
                {
                    p.WaitForExit(30000);
                }
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPdfPath))
                        File.Delete(tempPdfPath);
                }
                catch { }
            }
        }

        private string PrintTestPage(string printerName)
        {
            try
            {
                string computerName = Environment.MachineName;
                string userName = Environment.UserName;
                string ipAddress = GetLocalIPAddress();

                using (var ms = new MemoryStream())
                using (var bw = new BinaryWriter(ms))
                {
                    // Reset
                    bw.Write(new byte[] { 0x1B, 0x40 }); // ESC @

                    // Bold + double size
                    bw.Write(new byte[] { 0x1B, 0x45, 0x01 }); // ESC E 1 (bold on)
                    bw.Write(new byte[] { 0x1D, 0x21, 0x11 }); // GS ! 17 (double width & height)

                    bw.Write(Encoding.ASCII.GetBytes("=== RAWPRINTER TEST PAGE ===\n\n\n"));

                    // Reset size/bold
                    bw.Write(new byte[] { 0x1B, 0x45, 0x00 }); // ESC E 0 (bold off)
                    bw.Write(new byte[] { 0x1D, 0x21, 0x00 }); // Normal font

                    // Print details
                    bw.Write(Encoding.ASCII.GetBytes($"Date&Time : {DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}\n"));
                    bw.Write(Encoding.ASCII.GetBytes($"Printer   : {printerName}\n"));
                    bw.Write(Encoding.ASCII.GetBytes($"Computer  : {computerName}\n"));
                    bw.Write(Encoding.ASCII.GetBytes($"User      : {userName}\n"));
                    bw.Write(Encoding.ASCII.GetBytes($"IP        : {ipAddress}\n"));
                    bw.Write(Encoding.ASCII.GetBytes($"Version   : 1.0\n\n"));

                    bw.Write(Encoding.ASCII.GetBytes("RawPrinter setup successfully.\n"));
                    bw.Write(Encoding.ASCII.GetBytes("=============================\n\n"));

                    // Feed & Cut
                    bw.Write(new byte[] { 0x1B, 0x64, 0x02 }); // Feed 2 lines
                    bw.Write(new byte[] { 0x1D, 0x56, 0x42, 0x00 }); // Full cut (GS V B 0)

                    bw.Flush();
                    PrintRawBytes(printerName, ms.ToArray());
                }

                return "{\"status\":\"ok\",\"message\":\"RawPrinter setup successfully\"}";
            }
            catch (Exception ex)
            {
                return $"{{\"error\":\"{ex.Message}\"}}";
            }
        }

        private string GetLocalIPAddress()
        {
            string localIP = "Unknown";
            var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    localIP = ip.ToString();
                    break;
                }
            }
            return localIP;
        }

        private void WriteResponse(HttpListenerContext ctx, string content, string contentType)
        {
            ctx.Response.AppendHeader("Access-Control-Allow-Origin", "*");
            ctx.Response.AppendHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
            ctx.Response.AppendHeader("Access-Control-Allow-Headers", "Content-Type, Accept");

            ctx.Response.ContentType = contentType;
            var buffer = Encoding.UTF8.GetBytes(content);
            ctx.Response.ContentLength64 = buffer.Length;
            ctx.Response.OutputStream.Write(buffer, 0, buffer.Length);
        }

        // ---------- RAW PRINT ----------
        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true)]
        static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "ClosePrinter")]
        static extern bool ClosePrinter(IntPtr hPrinter);

        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true)]
        static extern bool StartDocPrinter(IntPtr hPrinter, int Level, [System.Runtime.InteropServices.In] ref DOCINFOA pDocInfo);

        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "EndDocPrinter")]
        static extern bool EndDocPrinter(IntPtr hPrinter);

        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "StartPagePrinter")]
        static extern bool StartPagePrinter(IntPtr hPrinter);

        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "EndPagePrinter")]
        static extern bool EndPagePrinter(IntPtr hPrinter);

        [System.Runtime.InteropServices.DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true)]
        static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct DOCINFOA
        {
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)]
            public string pDocName;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)]
            public string pOutputFile;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)]
            public string pDataType;
        }

        private void PrintRawBytes(string printerName, byte[] bytes)
        {
            IntPtr hPrinter;
            if (!OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
                throw new Exception("Could not open printer");

            var di = new DOCINFOA
            {
                pDocName = "RawPrintJob",
                pDataType = "RAW"
            };

            if (!StartDocPrinter(hPrinter, 1, ref di))
                throw new Exception("Could not start document");

            if (!StartPagePrinter(hPrinter))
                throw new Exception("Could not start page");

            IntPtr unmanagedPointer = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(bytes.Length);
            System.Runtime.InteropServices.Marshal.Copy(bytes, 0, unmanagedPointer, bytes.Length);

            int written;
            WritePrinter(hPrinter, unmanagedPointer, bytes.Length, out written);

            EndPagePrinter(hPrinter);
            EndDocPrinter(hPrinter);
            ClosePrinter(hPrinter);

            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(unmanagedPointer);

            WaitForPrinterQueueToEmpty(printerName);
        }

        private void WaitForPrinterQueueToEmpty(string printerName)
        {
            try
            {
                using (var searcher = new System.Management.ManagementObjectSearcher(
                    $"SELECT * FROM Win32_PrintJob WHERE Name LIKE '%{printerName}%'"))
                {
                    bool hasJobs;
                    do
                    {
                        hasJobs = false;
                        foreach (var job in searcher.Get())
                        {
                            hasJobs = true;
                            string status = job["Status"]?.ToString();
                            Console.WriteLine($"[INFO] Printer status: {status}");
                        }

                        if (hasJobs)
                            Thread.Sleep(1000);

                    } while (hasJobs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Printer status check failed: {ex.Message}");
            }
        }

        private string CheckPrinterStatus(string portName)
        {
            try
            {
                using (var serial = new SerialPort(portName, 9600, Parity.None, 8, StopBits.One))
                {
                    serial.ReadTimeout = 2000;
                    serial.Open();

                    byte[] statusCmd = new byte[] { 0x10, 0x04, 0x01 };
                    serial.Write(statusCmd, 0, statusCmd.Length);

                    int statusByte = serial.ReadByte();

                    if ((statusByte & 0x08) != 0) return "Paper end";
                    if ((statusByte & 0x20) != 0) return "Printer offline";

                    return "Printer ready";
                }
            }
            catch (Exception ex)
            {
                return $"Status check failed: {ex.Message}";
            }
        }

        private string GetPrinterStatus(string printerName)
        {
            try
            {
                using (var searcher = new System.Management.ManagementObjectSearcher(
                    "SELECT * FROM Win32_Printer WHERE Name = '" + printerName.Replace("\\", "\\\\") + "'"))
                {
                    foreach (var printer in searcher.Get())
                    {
                        var status = printer["PrinterStatus"];
                        var work = printer["WorkOffline"];

                        if (Convert.ToBoolean(work)) return "Offline";

                        switch (Convert.ToUInt16(status))
                        {
                            case 1: return "Other";
                            case 2: return "Unknown";
                            case 3: return "Idle";
                            case 4: return "Printing";
                            case 5: return "Warmup";
                            case 6: return "Stopped printing";
                            case 7: return "Offline";
                        }
                    }
                }

                return "Not found";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

    }

    static class Program
    {
        static void Main(string[] args)
        {
            if (Environment.UserInteractive)
            {
                Console.WriteLine("[INFO] Running RawPrintService in console mode...");
                var svc = new PrinterService();
                svc.GetType().GetMethod("OnStart", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                             .Invoke(svc, new object[] { args });

                Console.WriteLine("Press ENTER to stop...");
                Console.ReadLine();

                svc.GetType().GetMethod("OnStop", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                             .Invoke(svc, null);
            }
            else
            {
                ServiceBase.Run(new PrinterService());
            }
        }
    }
}