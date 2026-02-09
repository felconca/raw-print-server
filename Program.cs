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
                _listener.Prefixes.Add("http://+:9100/");
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
                // Handle preflight CORS
                if (ctx.Request.HttpMethod == "OPTIONS")
                {
                    ctx.Response.AppendHeader("Access-Control-Allow-Origin", "*");
                    ctx.Response.AppendHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
                    ctx.Response.AppendHeader("Access-Control-Allow-Headers", "Content-Type, Accept");
                    ctx.Response.StatusCode = 200;
                    ctx.Response.OutputStream.Close();
                    return;
                }

                if (ctx.Request.Url.AbsolutePath == "/")
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
                else if (ctx.Request.Url.AbsolutePath == "/print" && ctx.Request.HttpMethod == "POST")
                {
                    using (var reader = new System.IO.StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
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
                            {
                                bytes = Convert.FromBase64String(data);
                            }
                            else
                            {
                                bytes = Encoding.ASCII.GetBytes(data);
                            }

                            PrintRawBytes(printerName, bytes);
                            WriteResponse(ctx, "{\"status\":\"sent to printer\"}", "application/json");
                        }
                        catch (Exception ex)
                        {
                            WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
                        }
                    }
                }
                else if (ctx.Request.Url.AbsolutePath == "/status" && ctx.Request.HttpMethod == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();
                        var doc = JsonSerializer.Deserialize<Dictionary<string, string>>(body);

                        if (!doc.ContainsKey("printer"))
                        {
                            WriteResponse(ctx, "{\"error\":\"Missing printer name\"}", "application/json");
                            return;
                        }

                        var printerName = doc["printer"];
                        string status = GetPrinterStatus(printerName);

                        WriteResponse(ctx, $"{{\"status\":\"{status}\"}}", "application/json");
                    }
                }
                else if (ctx.Request.Url.AbsolutePath == "/test" && ctx.Request.HttpMethod == "POST")
                {
                    using (var reader = new StreamReader(ctx.Request.InputStream, ctx.Request.ContentEncoding))
                    {
                        var body = reader.ReadToEnd();
                        var doc = JsonSerializer.Deserialize<Dictionary<string, string>>(body);

                        if (!doc.ContainsKey("printer"))
                        {
                            WriteResponse(ctx, "{\"error\":\"Missing printer name\"}", "application/json");
                            return;
                        }

                        var printerName = doc["printer"];

                        try
                        {
                            string result = PrintTestPage(printerName);

                            // Proper JSON response
                            var json = JsonSerializer.Serialize(new
                            {
                                status = "ok",
                                message = result
                            });

                            WriteResponse(ctx, json, "application/json");
                        }
                        catch (Exception ex)
                        {
                            var json = JsonSerializer.Serialize(new
                            {
                                error = ex.Message
                            });
                            WriteResponse(ctx, json, "application/json");
                        }
                    }
                }

                else
                {
                    ctx.Response.StatusCode = 404;
                    WriteResponse(ctx, "{\"error\":\"Not Found\"}", "application/json");
                }
            }
            catch (Exception ex)
            {
                WriteResponse(ctx, $"{{\"error\":\"{ex.Message}\"}}", "application/json");
            }
            finally
            {
                ctx.Response.OutputStream.Close();
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

        // Helper to get local IPv4 address
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

            // 🔹 Poll spooler until the printer has no jobs left
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
                            Thread.Sleep(1000); // wait 1s before checking again

                    } while (hasJobs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR] Printer status check failed: {ex.Message}");
            }
        }


        // ---------- ESC/POS STATUS ----------
        private string CheckPrinterStatus(string portName)
        {
            try
            {
                using (var serial = new SerialPort(portName, 9600, Parity.None, 8, StopBits.One))
                {
                    serial.ReadTimeout = 2000;
                    serial.Open();

                    byte[] statusCmd = new byte[] { 0x10, 0x04, 0x01 }; // DLE EOT 1
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
