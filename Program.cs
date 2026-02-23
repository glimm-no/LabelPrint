using System.Buffers;
using System.Globalization;
using System.IO.Ports;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Streams;

internal static class Program
{
    private const double MmPerInch = 25.4;
    private const int FallbackDpi = 203;

    private static async Task<int> Main(string[] args)
    {
        if (!OperatingSystem.IsWindows())
        {
            Console.Error.WriteLine("This tool only runs on Windows.");
            return 1;
        }

        var options = LabelPrintOptions.Parse(args);
        if (options is null)
        {
            LabelPrintOptions.PrintUsage();
            return 1;
        }

        if (options.ListPrinters)
        {
            Console.WriteLine("=== Installed Printers (use with --printer) ===");
            foreach (var printer in PrinterUtilities.GetInstalledPrinters())
            {
                Console.WriteLine($"  {printer}");
            }

            Console.WriteLine();
            Console.WriteLine("=== COM Ports (use with --port for Bluetooth SPP) ===");
            foreach (var port in SerialPortSender.GetPortNames())
            {
                Console.WriteLine($"  {port}");
            }

            return 0;
        }

        if (options.ShowVersion)
        {
            Console.WriteLine(LabelPrintOptions.GetVersionString());
            return 0;
        }

        if (options.Install)
        {
            return VirtualPrinterInstaller.Install(
                options.VirtualPrinterName!,
                options.WatchFolder!,
                options.PrinterName,
                options.Port,
                options.Baud,
                options);
        }

        if (options.Uninstall)
        {
            return VirtualPrinterInstaller.Uninstall(options.VirtualPrinterName!);
        }

        if (options.RegisterStartup)
        {
            return StartupTaskRegistrar.Register(options);
        }

        if (options.UnregisterStartup)
        {
            return StartupTaskRegistrar.Unregister();
        }

        if (options.Watch)
        {
            return await WatchMode.RunAsync(options).ConfigureAwait(false);
        }

        if (!File.Exists(options.PdfPath))
        {
            Console.Error.WriteLine($"PDF file not found: {options.PdfPath}");
            return 1;
        }

        try
        {
            return await PrintPdfAsync(options).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Print failed: {ex.Message}");
            return 1;
        }
    }

    internal static async Task<int> PrintPdfAsync(LabelPrintOptions options)
    {
        var dpi = options.Dpi
            ?? (options.Port is null ? PrinterUtilities.TryGetPrinterDpi(options.PrinterName!) : null)
            ?? FallbackDpi;
        var resolvedOptions = options with { Dpi = dpi };
        var bitmap = await RenderPdfPageAsync(resolvedOptions).ConfigureAwait(false);
        var tspl = TsplBuilder.Build(resolvedOptions, bitmap);

        if (options.Port is not null)
        {
            SerialPortSender.Send(options.Port, options.Baud, tspl);
        }
        else
        {
            RawPrinterSender.Send(resolvedOptions.PrinterName!, tspl);
        }

        Console.WriteLine("Print job sent.");
        return 0;
    }

    private static async Task<BitmapData> RenderPdfPageAsync(LabelPrintOptions options)
    {
        var storageFile = await StorageFile.GetFileFromPathAsync(options.PdfPath);
        var pdf = await PdfDocument.LoadFromFileAsync(storageFile);

        if (options.PageIndex < 0 || options.PageIndex >= pdf.PageCount)
        {
            throw new ArgumentOutOfRangeException(nameof(options.PageIndex), $"PDF page index {options.PageIndex} is out of range.");
        }

        using var page = pdf.GetPage((uint)options.PageIndex);
        // Dpi is always resolved to a concrete value before RenderPdfPageAsync is called.
        var dpi = options.Dpi!.Value;
        var widthPx = (int)Math.Round(options.LabelWidthMm * dpi / MmPerInch);
        var heightPx = (int)Math.Round(options.LabelHeightMm * dpi / MmPerInch);

        var renderOptions = new PdfPageRenderOptions
        {
            DestinationWidth = (uint)widthPx,
            DestinationHeight = (uint)heightPx,
            // Construct white explicitly; Windows.UI.Colors static class requires the old SDK package.
            BackgroundColor = new Windows.UI.Color { A = 255, R = 255, G = 255, B = 255 },
        };

        using var stream = new InMemoryRandomAccessStream();
        await page.RenderToStreamAsync(stream, renderOptions);
        stream.Seek(0);

        var decoder = await BitmapDecoder.CreateAsync(stream);
        var pixelData = await decoder.GetPixelDataAsync(
            BitmapPixelFormat.Bgra8,
            BitmapAlphaMode.Premultiplied,
            new BitmapTransform(),
            ExifOrientationMode.IgnoreExifOrientation,
            ColorManagementMode.DoNotColorManage);

        var pixels = pixelData.DetachPixelData();
        return new BitmapData(widthPx, heightPx, pixels);
    }
}

internal sealed record LabelPrintOptions(
    string? PdfPath,
    string? PrinterName,
    double LabelWidthMm,
    double LabelHeightMm,
    int PageIndex,
    byte Threshold,
    int? Dpi,
    double? GapMm,
    double? GapOffsetMm,
    double? OffsetMm,
    int? Speed,
    int? Density,
    int? Direction,
    bool Tear,
    bool Peel,
    bool ListPrinters,
    bool ShowVersion,
    string? Port,
    int Baud,
    // Watch / install / startup modes
    bool Watch,
    string? WatchFolder,
    bool Install,
    bool Uninstall,
    string? VirtualPrinterName,
    bool RegisterStartup,
    bool UnregisterStartup,
    double? FeedMm,
    bool NoTear)
{
    public static LabelPrintOptions? Parse(string[] args)
    {
        if (args.Length == 0)
        {
            return null;
        }

        string? pdfPath = null;
        string? printerName = null;
        int? dpi = null;
        var widthMm = 100.0; // 10 cm
        var heightMm = 150.0; // 15 cm
        var pageIndex = 0;
        var threshold = (byte)180;
        double? gapMm = 3.0;
        double? gapOffsetMm = null;
        double? offsetMm = null;
        int? speed = null;
        int? density = null;
        int? direction = null;
        var tear = true;
        var peel = false;
        var listPrinters = false;
        var showVersion = false;
        string? port = null;
        var baud = 9600;
        var watch = false;
        string? watchFolder = null;
        var install = false;
        var uninstall = false;
        string? virtualPrinterName = "Label Printer";
        var registerStartup = false;
        var unregisterStartup = false;
        double? feedMm = null;
        var noTear = false;

        for (var i = 0; i < args.Length; i++)
        {
            var value = args[i];
            switch (value)
            {
                case "--list-printers":
                    listPrinters = true;
                    break;

                case "--version":
                    showVersion = true;
                    break;

                case "--pdf" when i + 1 < args.Length:
                    pdfPath = args[++i];
                    break;

                case "--printer" when i + 1 < args.Length:
                    printerName = args[++i];
                    break;

                case "--dpi" when i + 1 < args.Length:
                    if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedDpi) || parsedDpi <= 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --dpi.");
                        return null;
                    }

                    dpi = parsedDpi;
                    break;

                case "--width-mm" when i + 1 < args.Length:
                    if (!double.TryParse(args[++i], NumberStyles.Float, CultureInfo.InvariantCulture, out widthMm) || widthMm <= 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --width-mm.");
                        return null;
                    }

                    break;

                case "--height-mm" when i + 1 < args.Length:
                    if (!double.TryParse(args[++i], NumberStyles.Float, CultureInfo.InvariantCulture, out heightMm) || heightMm <= 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --height-mm.");
                        return null;
                    }

                    break;

                case "--page" when i + 1 < args.Length:
                    if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out pageIndex) || pageIndex < 1)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --page.");
                        return null;
                    }

                    pageIndex -= 1;
                    break;

                case "--threshold" when i + 1 < args.Length:
                    if (!byte.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out threshold))
                    {
                        Console.Error.WriteLine("Invalid value supplied to --threshold.");
                        return null;
                    }

                    break;

                case "--gap-mm" when i + 1 < args.Length:
                    if (!double.TryParse(args[++i], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedGap) || parsedGap < 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --gap-mm.");
                        return null;
                    }

                    gapMm = parsedGap;
                    break;

                case "--gap-offset-mm" when i + 1 < args.Length:
                    if (!double.TryParse(args[++i], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedGapOffset) || parsedGapOffset < 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --gap-offset-mm.");
                        return null;
                    }

                    gapOffsetMm = parsedGapOffset;
                    break;

                case "--offset-mm" when i + 1 < args.Length:
                    if (!double.TryParse(args[++i], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedOffset) || parsedOffset < 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --offset-mm.");
                        return null;
                    }

                    offsetMm = parsedOffset;
                    break;

                case "--speed" when i + 1 < args.Length:
                    if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedSpeed) || parsedSpeed <= 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --speed.");
                        return null;
                    }

                    speed = parsedSpeed;
                    break;

                case "--density" when i + 1 < args.Length:
                    if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedDensity) || parsedDensity <= 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --density.");
                        return null;
                    }

                    density = parsedDensity;
                    break;

                case "--direction" when i + 1 < args.Length:
                    if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedDirection) || (parsedDirection != 0 && parsedDirection != 1))
                    {
                        Console.Error.WriteLine("Invalid value supplied to --direction. Use 0 or 1.");
                        return null;
                    }

                    direction = parsedDirection;
                    break;

                case "--tear":
                    tear = true;
                    break;

                case "--no-tear":
                    noTear = true;
                    tear = false;
                    break;

                case "--peel":
                    peel = true;
                    break;

                case "--feed" when i + 1 < args.Length:
                    if (!double.TryParse(args[++i], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedFeed) || parsedFeed < 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --feed.");
                        return null;
                    }
                    feedMm = parsedFeed;
                    break;

                case "--watch":
                    watch = true;
                    break;

                case "--watch-folder" when i + 1 < args.Length:
                    watchFolder = args[++i];
                    break;

                case "--install":
                    install = true;
                    break;

                case "--uninstall":
                    uninstall = true;
                    break;

                case "--virtual-printer-name" when i + 1 < args.Length:
                    virtualPrinterName = args[++i];
                    break;

                case "--register-startup":
                    registerStartup = true;
                    break;

                case "--unregister-startup":
                    unregisterStartup = true;
                    break;

                case "--port" when i + 1 < args.Length:
                    port = args[++i];
                    break;

                case "--baud" when i + 1 < args.Length:
                    if (!int.TryParse(args[++i], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedBaud) || parsedBaud <= 0)
                    {
                        Console.Error.WriteLine("Invalid value supplied to --baud.");
                        return null;
                    }

                    baud = parsedBaud;
                    break;

                default:
                    Console.Error.WriteLine($"Unknown argument '{value}'.");
                    return null;
            }
        }

        if (listPrinters)
        {
            return new LabelPrintOptions(
                null, null,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, listPrinters, showVersion,
                port, baud,
                watch, watchFolder, install, uninstall, virtualPrinterName,
                registerStartup, unregisterStartup,
                feedMm, noTear);
        }

        if (showVersion)
        {
            return new LabelPrintOptions(
                null, null,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, false, true,
                port, baud,
                false, null, false, false, virtualPrinterName,
                false, false,
                feedMm, noTear);
        }

        if (unregisterStartup)
        {
            return new LabelPrintOptions(
                null, null,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, false, false,
                port, baud,
                false, null, false, false, virtualPrinterName,
                false, true,
                feedMm, noTear);
        }

        if (registerStartup)
        {
            if (string.IsNullOrWhiteSpace(watchFolder))
            {
                Console.Error.WriteLine("--register-startup requires --watch-folder.");
                return null;
            }

            if (string.IsNullOrWhiteSpace(printerName) && string.IsNullOrWhiteSpace(port))
            {
                Console.Error.WriteLine("--register-startup requires --printer or --port.");
                return null;
            }

            return new LabelPrintOptions(
                null, printerName,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, false, false,
                port, baud,
                false, watchFolder, false, false, virtualPrinterName,
                true, false,
                feedMm, noTear);
        }

        if (install)
        {
            if (string.IsNullOrWhiteSpace(watchFolder))
            {
                Console.Error.WriteLine("--install requires --watch-folder.");
                return null;
            }

            if (string.IsNullOrWhiteSpace(printerName) && string.IsNullOrWhiteSpace(port))
            {
                Console.Error.WriteLine("--install requires --printer or --port (the physical label printer).");
                return null;
            }

            return new LabelPrintOptions(
                null, printerName,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, false, false,
                port, baud,
                false, watchFolder, install, false, virtualPrinterName,
                false, false,
                feedMm, noTear);
        }

        if (uninstall)
        {
            return new LabelPrintOptions(
                null, null,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, false, false,
                port, baud,
                false, null, false, true, virtualPrinterName,
                false, false,
                feedMm, noTear);
        }

        if (watch)
        {
            if (string.IsNullOrWhiteSpace(watchFolder))
            {
                Console.Error.WriteLine("--watch requires --watch-folder.");
                return null;
            }

            if (string.IsNullOrWhiteSpace(printerName) && string.IsNullOrWhiteSpace(port))
            {
                Console.Error.WriteLine("--watch requires --printer or --port (the physical label printer).");
                return null;
            }

            return new LabelPrintOptions(
                null, printerName,
                widthMm, heightMm, pageIndex, threshold,
                dpi, gapMm, gapOffsetMm, offsetMm,
                speed, density, direction,
                tear, peel, false, false,
                port, baud,
                true, watchFolder, false, false, virtualPrinterName,
                false, false,
                feedMm, noTear);
        }

        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(printerName) && string.IsNullOrWhiteSpace(port))
        {
            Console.Error.WriteLine("Either --printer or --port must be specified.");
            return null;
        }

        if (!string.IsNullOrWhiteSpace(printerName) && !string.IsNullOrWhiteSpace(port))
        {
            Console.Error.WriteLine("--printer and --port cannot be used together.");
            return null;
        }

        return new LabelPrintOptions(
            pdfPath,
            printerName,
            widthMm, heightMm, pageIndex, threshold,
            dpi, gapMm, gapOffsetMm, offsetMm,
            speed, density, direction,
            tear, peel, listPrinters, false,
            port, baud,
            false, null, false, false, null,
            false, false,
            feedMm, noTear);
    }

    public static string GetVersionString()
    {
        var assembly = typeof(LabelPrintOptions).Assembly;
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        var version = !string.IsNullOrWhiteSpace(informationalVersion)
            ? informationalVersion
            : assembly.GetName().Version?.ToString() ?? "unknown";

        return $"LabelPrint {version}";
    }

    public static void PrintUsage()
    {
        Console.WriteLine(GetVersionString());
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  LabelPrint --pdf <path> [--printer <name> | --port <COMn>] [options]  Print a single PDF");
        Console.WriteLine("  LabelPrint --watch --watch-folder <dir> [--printer <name> | --port <COMn>] [options]  Watch folder mode");
        Console.WriteLine("  LabelPrint --install --watch-folder <dir> [--printer <name> | --port <COMn>] [options]  Install virtual printer (default: 'Label Printer') (run as admin)");
        Console.WriteLine("  LabelPrint --uninstall [--virtual-printer-name <name>]      Remove virtual printer (default: 'Label Printer') (run as admin)");
        Console.WriteLine("  LabelPrint --register-startup --watch-folder <dir> [--printer <name> | --port <COMn>] [options]  Register as logon task");
        Console.WriteLine("  LabelPrint --unregister-startup                             Remove logon task");
        Console.WriteLine("  LabelPrint --list-printers                                  List printers and COM ports");
        Console.WriteLine("  LabelPrint --version                                        Print version");
        Console.WriteLine();
        Console.WriteLine("Transport (one required for print/watch/install):");
        Console.WriteLine("  --printer <name>              Windows printer name (use --list-printers to see names).");
        Console.WriteLine("  --port <COMn>                 COM port for Bluetooth SPP printers (e.g. COM4).");
        Console.WriteLine("  --baud <n>                    Baud rate for --port (default 9600).");
        Console.WriteLine();
        Console.WriteLine("Watch / install options:");
        Console.WriteLine("  --watch-folder <dir>          Folder to watch for new PDF files.");
        Console.WriteLine("  --virtual-printer-name <name> Name of the virtual printer (default: 'Label Printer').");
        Console.WriteLine();
        Console.WriteLine("Label options:");
        Console.WriteLine("  --dpi <n>                     Render DPI (defaults to printer capability or 203).");
        Console.WriteLine("  --width-mm <n>                Label width in mm (default 100).");
        Console.WriteLine("  --height-mm <n>               Label height in mm (default 150).");
        Console.WriteLine("  --page <n>                    1-based page index (default 1).");
        Console.WriteLine("  --threshold <n>               0-255 threshold for black/white (default 180).");
        Console.WriteLine("  --gap-mm <n>                  Gap in mm (defaults to printer setting).");
        Console.WriteLine("  --gap-offset-mm <n>           Gap offset in mm (defaults to printer setting).");
        Console.WriteLine("  --offset-mm <n>               Vertical label offset in mm (defaults to printer setting).");
        Console.WriteLine("  --speed <n>                   Print speed (TSPL SPEED).");
        Console.WriteLine("  --density <n>                 Print density (TSPL DENSITY).");
        Console.WriteLine("  --direction <n>               Print direction 0 or 1 (defaults to printer setting).");
        Console.WriteLine("  --tear                        Enable tear mode (default).");
        Console.WriteLine("  --no-tear                     Disable tear mode.");
        Console.WriteLine("  --peel                        Enable peel mode.");
        Console.WriteLine("  --feed <n>                    Feed n mm after print.");
    }
}

internal sealed record BitmapData(int Width, int Height, byte[] Pixels);

internal static class TsplBuilder
{
    public static byte[] Build(LabelPrintOptions options, BitmapData bitmap)
    {
        var widthBytes = (bitmap.Width + 7) / 8;
        var dataSize = widthBytes * bitmap.Height;
        var raster = ArrayPool<byte>.Shared.Rent(dataSize);
        try
        {
            Array.Clear(raster, 0, dataSize);
            FillRaster(options, bitmap, raster, widthBytes);

            // Fix inverted output
            InvertRaster(raster, dataSize);

            var header = new StringBuilder();
            // SIZE uses integer mm to match vendor driver behaviour (LABELCORE.dll uses %d format).
            header.Append(CultureInfo.InvariantCulture, $"SIZE {(int)options.LabelWidthMm} mm,{(int)options.LabelHeightMm} mm\r\n");
            if (options.GapMm is not null || options.GapOffsetMm is not null)
            {
                // GAP uses integer mm to match vendor driver behaviour.
                var gap = (int)(options.GapMm ?? 0.0);
                var gapOffset = (int)(options.GapOffsetMm ?? 0.0);
                header.Append(CultureInfo.InvariantCulture, $"GAP {gap} mm,{gapOffset} mm\r\n");
            }
            if (options.OffsetMm is not null)
            {
                header.Append(CultureInfo.InvariantCulture, $"OFFSET {(int)options.OffsetMm.Value} mm\r\n");
            }
            if (options.Direction is not null)
            {
                header.Append(CultureInfo.InvariantCulture, $"DIRECTION {options.Direction.Value}\r\n");
            }
            if (options.Speed is not null)
            {
                header.Append(CultureInfo.InvariantCulture, $"SPEED {options.Speed.Value}\r\n");
            }
            if (options.Density is not null)
            {
                header.Append(CultureInfo.InvariantCulture, $"DENSITY {options.Density.Value}\r\n");
            }
            if (options.Tear)
            {
                header.Append("SET TEAR ON\r\n");
            }
            if (options.Peel)
            {
                header.Append("SET PEEL ON\r\n");
            }

            if (options.FeedMm is not null)
            {
                header.Append(CultureInfo.InvariantCulture, $"FEED {(int)options.FeedMm.Value} mm\r\n");
            }

            header.Append("CLS\r\n");
            header.Append(CultureInfo.InvariantCulture, $"BITMAP 0,0,{widthBytes},{bitmap.Height},0,");

            var footer = "\r\nPRINT 1,1\r\n";

            var headerBytes = Encoding.ASCII.GetBytes(header.ToString());
            var footerBytes = Encoding.ASCII.GetBytes(footer);

            var payload = new byte[headerBytes.Length + dataSize + footerBytes.Length];
            // Qualify System.Buffer to avoid ambiguity with Windows.Storage.Streams.Buffer.
            System.Buffer.BlockCopy(headerBytes, 0, payload, 0, headerBytes.Length);
            System.Buffer.BlockCopy(raster, 0, payload, headerBytes.Length, dataSize);
            System.Buffer.BlockCopy(footerBytes, 0, payload, headerBytes.Length + dataSize, footerBytes.Length);
            return payload;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(raster);
        }
    }

    private static void FillRaster(LabelPrintOptions options, BitmapData bitmap, byte[] raster, int widthBytes)
    {
        var threshold = options.Threshold;
        var pixels = bitmap.Pixels;
        var stride = bitmap.Width * 4;

        for (var y = 0; y < bitmap.Height; y++)
        {
            var rowOffset = y * stride;
            var outOffset = y * widthBytes;
            for (var x = 0; x < bitmap.Width; x++)
            {
                var pixelOffset = rowOffset + (x * 4);
                var b = pixels[pixelOffset];
                var g = pixels[pixelOffset + 1];
                var r = pixels[pixelOffset + 2];

                // Integer approximation of BT.601 luma (equivalent to float but ~3x faster in a tight loop).
                var luminance = (byte)((r * 299 + g * 587 + b * 114) / 1000);
                var isBlack = luminance < threshold;
                if (!isBlack)
                {
                    continue;
                }

                var byteIndex = outOffset + (x / 8);
                var bitIndex = 7 - (x % 8);
                raster[byteIndex] |= (byte)(1 << bitIndex);
            }
        }
    }

    private static void InvertRaster(byte[] raster, int length)
    {
        for (var i = 0; i < length; i++)
        {
            raster[i] = (byte)~raster[i];
        }
    }
}

internal static class RawPrinterSender
{
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int StartDocPrinter(IntPtr hPrinter, int level, [In] ref DOC_INFO_1 docInfo);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool WritePrinter(IntPtr hPrinter, byte[] pBytes, int dwCount, out int dwWritten);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct DOC_INFO_1
    {
        public string pDocName;
        public string pOutputFile;
        public string pDatatype;
    }

    public static void Send(string printerName, byte[] payload)
    {
        if (!OpenPrinter(printerName, out var printerHandle, IntPtr.Zero))
        {
            throw new InvalidOperationException($"OpenPrinter failed: {Marshal.GetLastWin32Error()}");
        }

        try
        {
            var docInfo = new DOC_INFO_1
            {
                pDocName = "Label Print",
                pOutputFile = string.Empty,
                pDatatype = "RAW",
            };

            if (StartDocPrinter(printerHandle, 1, ref docInfo) == 0)
            {
                throw new InvalidOperationException($"StartDocPrinter failed: {Marshal.GetLastWin32Error()}");
            }

            try
            {
                if (!StartPagePrinter(printerHandle))
                {
                    throw new InvalidOperationException($"StartPagePrinter failed: {Marshal.GetLastWin32Error()}");
                }

                try
                {
                    if (!WritePrinter(printerHandle, payload, payload.Length, out var written) || written != payload.Length)
                    {
                        throw new InvalidOperationException($"WritePrinter failed: {Marshal.GetLastWin32Error()}");
                    }
                }
                finally
                {
                    EndPagePrinter(printerHandle);
                }
            }
            finally
            {
                EndDocPrinter(printerHandle);
            }
        }
        finally
        {
            ClosePrinter(printerHandle);
        }
    }
}

internal static class PrinterUtilities
{
    private const int PrinterEnumLocal = 0x00000002;
    private const int PrinterEnumConnections = 0x00000004;
    private const int ErrorInsufficientBuffer = 122;
    private const int LOGPIXELSX = 88;
    private const int LOGPIXELSY = 90;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PRINTER_INFO_4
    {
        public string pPrinterName;
        public string pServerName;
        public uint Attributes;
    }

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool EnumPrinters(int flags, string? name, int level, IntPtr pPrinterEnum, int cbBuf, out int pcbNeeded, out int pcReturned);

    [DllImport("gdi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateDC(string lpszDriver, string lpszDevice, string? lpszOutput, IntPtr lpInitData);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    [DllImport("gdi32.dll", SetLastError = true)]
    private static extern bool DeleteDC(IntPtr hdc);

    public static IEnumerable<string> GetInstalledPrinters()
    {
        var flags = PrinterEnumLocal | PrinterEnumConnections;
        if (!EnumPrinters(flags, null, 4, IntPtr.Zero, 0, out var needed, out _))
        {
            var error = Marshal.GetLastWin32Error();
            if (error != ErrorInsufficientBuffer)
            {
                yield break;
            }
        }

        if (needed <= 0)
        {
            yield break;
        }

        var buffer = Marshal.AllocHGlobal(needed);
        try
        {
            if (!EnumPrinters(flags, null, 4, buffer, needed, out _, out var returned))
            {
                yield break;
            }

            var offset = buffer;
            var structSize = Marshal.SizeOf<PRINTER_INFO_4>();
            for (var i = 0; i < returned; i++)
            {
                var info = Marshal.PtrToStructure<PRINTER_INFO_4>(offset);
                if (!string.IsNullOrWhiteSpace(info.pPrinterName))
                {
                    yield return info.pPrinterName;
                }

                offset = IntPtr.Add(offset, structSize);
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    public static int? TryGetPrinterDpi(string printerName)
    {
        var hdc = CreateDC("WINSPOOL", printerName, null, IntPtr.Zero);
        if (hdc == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            var x = GetDeviceCaps(hdc, LOGPIXELSX);
            var y = GetDeviceCaps(hdc, LOGPIXELSY);
            if (x <= 0 && y <= 0)
            {
                return null;
            }

            if (x > 0 && y > 0)
            {
                return (x + y) / 2;
            }

            return x > 0 ? x : y;
        }
        finally
        {
            DeleteDC(hdc);
        }
    }
}

/// <summary>
/// Sends raw TSPL bytes directly to a serial COM port (Bluetooth SPP profile).
/// Use when the printer appears as a virtual COM port rather than a Windows printer.
/// </summary>
internal static class SerialPortSender
{
    /// <summary>Returns the names of all available serial COM ports on this machine.</summary>
    public static IEnumerable<string> GetPortNames() => SerialPort.GetPortNames().OrderBy(p => p);

    /// <summary>
    /// Opens <paramref name="portName"/> at <paramref name="baud"/> baud (8N1, no flow control),
    /// writes <paramref name="payload"/>, and closes the port.
    /// </summary>
    public static void Send(string portName, int baud, byte[] payload)
    {
        using var port = new SerialPort(portName, baud, Parity.None, 8, StopBits.One)
        {
            Handshake = Handshake.None,
            WriteTimeout = 30_000,
        };

        port.Open();
        port.BaseStream.Write(payload, 0, payload.Length);
        port.BaseStream.Flush();
    }
}

/// <summary>
/// Watches a folder for new PDF files and prints each one automatically.
/// Run with Ctrl+C to stop.
/// </summary>
internal static class WatchMode
{
    // How long to wait after a file appears before trying to open it (gives the writer time to finish).
    private static readonly TimeSpan SettleDelay = TimeSpan.FromMilliseconds(800);

    public static async Task<int> RunAsync(LabelPrintOptions options)
    {
        var folder = options.WatchFolder!;
        Directory.CreateDirectory(folder);

        Console.WriteLine($"Watching '{folder}' for PDF files. Press Ctrl+C to stop.");
        Console.WriteLine($"Printing to: {(options.Port is not null ? $"COM port {options.Port}" : options.PrinterName)}");

        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            cts.Cancel();
        };

        // Process any PDFs already sitting in the folder at startup.
        foreach (var existing in Directory.GetFiles(folder, "*.pdf"))
        {
            await ProcessFileAsync(existing, options).ConfigureAwait(false);
        }

        using var watcher = new FileSystemWatcher(folder, "*.pdf")
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
            EnableRaisingEvents = true,
        };

        // Queue of files to process — avoids re-entrancy from rapid events.
        var queue = new System.Collections.Concurrent.ConcurrentQueue<string>();
        watcher.Created += (_, e) => queue.Enqueue(e.FullPath);
        watcher.Renamed += (_, e) => queue.Enqueue(e.FullPath);

        while (!cts.Token.IsCancellationRequested)
        {
            if (queue.TryDequeue(out var path))
            {
                // Brief settle delay so the writer has finished before we open the file.
                await Task.Delay(SettleDelay, cts.Token).ConfigureAwait(false);
                await ProcessFileAsync(path, options).ConfigureAwait(false);
            }
            else
            {
                await Task.Delay(100, cts.Token).ConfigureAwait(false);
            }
        }

        Console.WriteLine("Watch mode stopped.");
        return 0;
    }

    private static async Task ProcessFileAsync(string path, LabelPrintOptions options)
    {
        if (!File.Exists(path))
        {
            return;
        }

        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Detected: {Path.GetFileName(path)}");
        try
        {
            var printOptions = options with { PdfPath = path };
            await Program.PrintPdfAsync(printOptions).ConfigureAwait(false);

            // Delete after successful print so the folder stays clean.
            try { File.Delete(path); } catch { /* best-effort */ }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[{DateTime.Now:HH:mm:ss}] ERROR printing {Path.GetFileName(path)}: {ex.Message}");
            // Move to an error subfolder so it doesn't get retried in an infinite loop.
            var errorDir = Path.Combine(Path.GetDirectoryName(path)!, "errors");
            Directory.CreateDirectory(errorDir);
            var dest = Path.Combine(errorDir, Path.GetFileName(path));
            try { File.Move(path, dest, overwrite: true); } catch { /* best-effort */ }
        }
    }
}

/// <summary>
/// Installs / uninstalls a silent virtual "Label Printer" in Windows that auto-saves
/// print jobs as PDF files to a fixed watch folder — no Save dialog shown to the user.
///
/// How it works:
///   1. Creates a "Local Port" named after the output file path (e.g. C:\LabelPrintQueue\label.pdf).
///      The built-in "Local Port" monitor treats a path as a FILE: port and writes directly to it.
///   2. Installs a printer using the "Microsoft Print To PDF" driver pointing at that port.
///   3. Stores the label options in a companion settings file next to the exe so --watch can
///      pick them up without the user having to re-specify them every time.
///
/// Requires elevation (admin rights) because AddPort and AddPrinter write to HKLM.
/// </summary>
internal static class VirtualPrinterInstaller
{
    // The "Microsoft Print To PDF" driver name as it appears in Windows.
    private const string PdfDriverName = "Microsoft Print To PDF";
    private const uint PrinterAttributeQueued = 0x00000001;
    private const uint PrinterAttributeLocal = 0x00000040;

    // Settings file written alongside the exe so --watch can read the config back.
    private static string SettingsPath =>
        Path.Combine(AppContext.BaseDirectory, "labelprint-watch.json");

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PORT_INFO_1
    {
        public IntPtr pName;
    }

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool AddPortEx(string? pName, uint level, IntPtr lpBuffer, string lpMonitorName);

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool DeletePort(string? pName, IntPtr hWnd, string pPortName);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PRINTER_INFO_2
    {
        public string? pServerName;
        public string pPrinterName;
        public string? pShareName;
        public string pPortName;
        public string pDriverName;
        public string? pComment;
        public string? pLocation;
        public IntPtr pDevMode;
        public string? pSepFile;
        public string pPrintProcessor;
        public string pDatatype;
        public string? pParameters;
        public IntPtr pSecurityDescriptor;
        public uint Attributes;
        public uint Priority;
        public uint DefaultPriority;
        public uint StartTime;
        public uint UntilTime;
        public uint Status;
        public uint cJobs;
        public uint AveragePPM;
    }

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr AddPrinter(string? pName, uint level, ref PRINTER_INFO_2 pPrinter);

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PRINTER_DEFAULTS
    {
        public string? pDatatype;
        public IntPtr pDevMode;
        public uint DesiredAccess;
    }

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, ref PRINTER_DEFAULTS pDefault);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", SetLastError = true)]
    private static extern bool DeletePrinter(IntPtr hPrinter);

    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern uint SetPrinterDataEx(
        IntPtr hPrinter,
        string pKeyName,
        string pValueName,
        uint Type,
        byte[] pData,
        uint cbData);

    private const uint RegSz = 1;
    private const uint RegDword = 4;
    private const uint ErrorSuccess = 0;
    private const uint PrinterAccessAdminister = 0x00000004;

    public static int Install(
        string virtualPrinterName,
        string watchFolder,
        string? physicalPrinter,
        string? comPort,
        int baud,
        LabelPrintOptions options)
    {
        try
        {
            var resolvedWatchFolder = Path.GetFullPath(watchFolder);
            if (resolvedWatchFolder.StartsWith(@"\\", StringComparison.Ordinal))
            {
                Console.Error.WriteLine("Install failed: --watch-folder must be a local path, not a UNC/network path.");
                return 1;
            }

            Directory.CreateDirectory(resolvedWatchFolder);
            EnsureSystemWriteAccess(resolvedWatchFolder);

            // The output file path used as the port name.
            // "Local Port" monitor accepts a full file path as a port name and writes directly to it.
            var outputFile = Path.Combine(resolvedWatchFolder, "label.pdf");

            Console.WriteLine($"Creating port '{outputFile}'...");
            AddLocalPort(outputFile);

            Console.WriteLine($"Installing printer '{virtualPrinterName}'...");
            var printerInfo = new PRINTER_INFO_2
            {
                pPrinterName = virtualPrinterName,
                pPortName = outputFile,
                pDriverName = PdfDriverName,
                pPrintProcessor = "winprint",
                pDatatype = "NT EMF 1.008",
                Attributes = PrinterAttributeLocal | PrinterAttributeQueued,
            };

            var handle = AddPrinter(null, 2, ref printerInfo);
            if (handle == IntPtr.Zero)
            {
                var err = Marshal.GetLastWin32Error();
                if (err == 0x57) // ERROR_INVALID_PARAMETER — driver not found
                {
                    Console.Error.WriteLine($"ERROR: The '{PdfDriverName}' driver is not installed on this machine.");
                    Console.Error.WriteLine("Enable it via: Settings → Bluetooth & devices → Printers & scanners → Add a printer → Microsoft Print to PDF.");
                    return 1;
                }

                Console.Error.WriteLine($"AddPrinter failed: Win32 error {err}");
                return 1;
            }

            ConfigureMicrosoftPrintToPdfOutput(handle, outputFile);
            ClosePrinter(handle);

            // Persist the watch settings so --watch can be run without re-specifying all options.
            SaveSettings(virtualPrinterName, resolvedWatchFolder, physicalPrinter, comPort, baud, options);

            Console.WriteLine();
            Console.WriteLine($"✓ Virtual printer '{virtualPrinterName}' installed.");
            Console.WriteLine($"  Output folder : {resolvedWatchFolder}");
            Console.WriteLine($"  Physical target: {(comPort is not null ? $"COM port {comPort}" : physicalPrinter)}");
            Console.WriteLine();
            Console.WriteLine("Next step — start the watch service:");
            Console.WriteLine($"  LabelPrint --watch --watch-folder \"{resolvedWatchFolder}\" --printer \"{physicalPrinter}\" --width-mm {(int)options.LabelWidthMm} --height-mm {(int)options.LabelHeightMm}");
            Console.WriteLine();
            Console.WriteLine("Or simply run:  LabelPrint --watch  (settings are saved)");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Install failed: {ex.Message}");
            return 1;
        }
    }

    public static int Uninstall(string virtualPrinterName)
    {
        try
        {
            var defaults = new PRINTER_DEFAULTS
            {
                DesiredAccess = PrinterAccessAdminister,
            };

            if (!OpenPrinter(virtualPrinterName, out var handle, ref defaults))
            {
                var openErr = Marshal.GetLastWin32Error();
                if (openErr == 1801) // ERROR_INVALID_PRINTER_NAME
                {
                    Console.WriteLine($"Printer '{virtualPrinterName}' is already removed.");
                    return 0;
                }

                Console.Error.WriteLine($"OpenPrinter failed for '{virtualPrinterName}': Win32 error {openErr}.");
                return 1;
            }

            if (!DeletePrinter(handle))
            {
                var deleteErr = Marshal.GetLastWin32Error();
                ClosePrinter(handle);
                Console.Error.WriteLine($"DeletePrinter failed: Win32 error {deleteErr}");
                return 1;
            }

            ClosePrinter(handle);

            // Remove the saved settings file if present.
            if (File.Exists(SettingsPath))
            {
                File.Delete(SettingsPath);
            }

            Console.WriteLine($"✓ Virtual printer '{virtualPrinterName}' removed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Uninstall failed: {ex.Message}");
            return 1;
        }
    }

    /// <summary>
    /// Adds a "Local Port" via the spooler API so AddPrinter recognizes it immediately.
    /// </summary>
    private static void AddLocalPort(string portName)
    {
        var portNamePtr = Marshal.StringToHGlobalUni(portName);
        var portInfo = new PORT_INFO_1 { pName = portNamePtr };
        var buffer = Marshal.AllocHGlobal(Marshal.SizeOf<PORT_INFO_1>());
        try
        {
            Marshal.StructureToPtr(portInfo, buffer, fDeleteOld: false);
            if (AddPortEx(null, 1, buffer, "Local Port"))
            {
                return;
            }

            var err = Marshal.GetLastWin32Error();
            if (err == 183) // ERROR_ALREADY_EXISTS
            {
                return;
            }

            if (err == 87) // ERROR_INVALID_PARAMETER
            {
                // Fallback for systems where AddPortEx with Local Port is rejected:
                // write the port entry where the spooler expects local monitor ports.
                AddLocalPortViaRegistry(portName);
                return;
            }

            throw new InvalidOperationException($"AddPortEx failed: Win32 error {err}");
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
            Marshal.FreeHGlobal(portNamePtr);
        }
    }

    private static void AddLocalPortViaRegistry(string portName)
    {
        var candidatePaths = new[]
        {
            @"SYSTEM\CurrentControlSet\Control\Print\Monitors\Local Port\Ports",
            @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports",
        };

        Exception? lastError = null;
        foreach (var path in candidatePaths)
        {
            try
            {
                using var portsKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(path, writable: true)
                    ?? Microsoft.Win32.Registry.LocalMachine.CreateSubKey(path, writable: true);
                if (portsKey is null)
                {
                    continue;
                }

                portsKey.SetValue(portName, string.Empty, Microsoft.Win32.RegistryValueKind.String);
                return;
            }
            catch (Exception ex)
            {
                lastError = ex;
            }
        }

        var details = lastError is null ? "no writable ports registry key was found" : lastError.Message;
        throw new InvalidOperationException($"Cannot create local port registry entry ({details}).");
    }

    private static void EnsureSystemWriteAccess(string folderPath)
    {
        var directoryInfo = new DirectoryInfo(folderPath);
        var security = directoryInfo.GetAccessControl();
        var systemSid = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
        var rule = new FileSystemAccessRule(
            systemSid,
            FileSystemRights.Modify | FileSystemRights.Synchronize,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
            PropagationFlags.None,
            AccessControlType.Allow);

        var changed = false;
        security.ModifyAccessRule(AccessControlModification.Add, rule, out changed);
        if (changed)
        {
            directoryInfo.SetAccessControl(security);
        }
    }

    /// <summary>
    /// Configures Microsoft Print To PDF to write to a fixed file without prompting.
    /// </summary>
    private static void ConfigureMicrosoftPrintToPdfOutput(IntPtr printerHandle, string outputFile)
    {
        // REG_SZ must include a UTF-16 null terminator.
        var outputBytes = Encoding.Unicode.GetBytes(outputFile + '\0');
        var outputResult = SetPrinterDataEx(
            printerHandle,
            "PrinterDriverData",
            "OutputFile",
            RegSz,
            outputBytes,
            (uint)outputBytes.Length);

        if (outputResult != ErrorSuccess)
        {
            throw new InvalidOperationException($"SetPrinterDataEx(OutputFile) failed: Win32 error {outputResult}");
        }

        // DWORD 0 = do not prompt for a file name.
        var promptBytes = BitConverter.GetBytes(0u);
        var promptResult = SetPrinterDataEx(
            printerHandle,
            "PrinterDriverData",
            "PromptForFileName",
            RegDword,
            promptBytes,
            (uint)promptBytes.Length);

        if (promptResult != ErrorSuccess)
        {
            throw new InvalidOperationException($"SetPrinterDataEx(PromptForFileName) failed: Win32 error {promptResult}");
        }
    }

    private static void SaveSettings(
        string virtualPrinterName,
        string watchFolder,
        string? physicalPrinter,
        string? comPort,
        int baud,
        LabelPrintOptions options)
    {
        // Simple JSON-like key=value file — avoids taking a JSON library dependency.
        var lines = new[]
        {
            $"virtualPrinterName={virtualPrinterName}",
            $"watchFolder={watchFolder}",
            $"physicalPrinter={physicalPrinter ?? string.Empty}",
            $"comPort={comPort ?? string.Empty}",
            $"baud={baud}",
            $"widthMm={options.LabelWidthMm.ToString(CultureInfo.InvariantCulture)}",
            $"heightMm={options.LabelHeightMm.ToString(CultureInfo.InvariantCulture)}",
            $"dpi={options.Dpi?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"threshold={options.Threshold}",
            $"gapMm={options.GapMm?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"gapOffsetMm={options.GapOffsetMm?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"offsetMm={options.OffsetMm?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"speed={options.Speed?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"density={options.Density?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"direction={options.Direction?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"tear={options.Tear}",
            $"peel={options.Peel}",
            $"feedMm={options.FeedMm?.ToString(CultureInfo.InvariantCulture) ?? string.Empty}",
            $"noTear={options.NoTear}",
        };

        File.WriteAllLines(SettingsPath, lines);
    }
}

/// <summary>
/// Registers / unregisters a Windows Task Scheduler logon task that starts the
/// LabelPrint watcher automatically when the current user logs in.
///
/// Uses <c>schtasks.exe</c> (built into every Windows version) so no extra packages
/// or admin rights are required — tasks registered for the current user run in the
/// user session and can access the desktop printer queue normally.
/// </summary>
internal static class StartupTaskRegistrar
{
    private const string TaskName = "LabelPrintWatcher";

    public static int Register(LabelPrintOptions options)
    {
        try
        {
            var exePath = Environment.ProcessPath
                ?? throw new InvalidOperationException("Cannot determine the path of the current executable.");

            // Build the argument string that will be passed to the exe at logon.
            var sb = new StringBuilder();
            sb.Append("--watch");
            sb.Append($" --watch-folder \"{options.WatchFolder}\"");

            if (!string.IsNullOrWhiteSpace(options.PrinterName))
            {
                sb.Append($" --printer \"{options.PrinterName}\"");
            }

            if (!string.IsNullOrWhiteSpace(options.Port))
            {
                sb.Append($" --port {options.Port}");
                sb.Append($" --baud {options.Baud}");
            }

            sb.Append(CultureInfo.InvariantCulture, $" --width-mm {options.LabelWidthMm}");
            sb.Append(CultureInfo.InvariantCulture, $" --height-mm {options.LabelHeightMm}");
            sb.Append(CultureInfo.InvariantCulture, $" --threshold {options.Threshold}");

            if (options.Dpi is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --dpi {options.Dpi.Value}");
            }

            if (options.GapMm is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --gap-mm {options.GapMm.Value}");
            }

            if (options.GapOffsetMm is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --gap-offset-mm {options.GapOffsetMm.Value}");
            }

            if (options.OffsetMm is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --offset-mm {options.OffsetMm.Value}");
            }

            if (options.Speed is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --speed {options.Speed.Value}");
            }

            if (options.Density is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --density {options.Density.Value}");
            }

            if (options.Direction is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" --direction {options.Direction.Value}");
            }

            if (options.Tear) sb.Append(" --tear");
            if (options.NoTear) sb.Append(" --no-tear");
            if (options.Peel) sb.Append(" --peel");
            if (options.FeedMm is not null) sb.Append(CultureInfo.InvariantCulture, $" --feed {options.FeedMm.Value}");

            var arguments = sb.ToString();

            // Delete any existing task with the same name first (idempotent re-registration).
            RunSchtasks($"/Delete /TN \"{TaskName}\" /F");

            // Create a new logon trigger task for the current user.
            // /SC ONLOGON  — trigger: at logon of the current user
            // /DELAY       — 1-minute delay (HH:MM format) so the spooler/Bluetooth stack are ready
            // /F           — force overwrite if it already exists
            // Note: /RL HIGHEST is intentionally omitted — it requires admin rights and the
            //       watcher does not need elevated privileges.
            var createArgs = $"/Create /TN \"{TaskName}\" /TR \"\\\"{exePath}\\\" {arguments}\" " +
                             $"/SC ONLOGON /DELAY 0000:01 /F";

            RunSchtasks(createArgs);

            Console.WriteLine($"✓ Startup task '{TaskName}' registered.");
            Console.WriteLine($"  Exe  : {exePath}");
            Console.WriteLine($"  Args : {arguments}");
            Console.WriteLine();
            Console.WriteLine("The watcher will start automatically the next time you log in.");
            Console.WriteLine("To start it now without logging out, run:");
            Console.WriteLine($"  LabelPrint --watch --watch-folder \"{options.WatchFolder}\" --printer \"{options.PrinterName}\"");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to register startup task: {ex.Message}");
            return 1;
        }
    }

    public static int Unregister()
    {
        try
        {
            RunSchtasks($"/Delete /TN \"{TaskName}\" /F");
            Console.WriteLine($"✓ Startup task '{TaskName}' removed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to remove startup task: {ex.Message}");
            return 1;
        }
    }

    private static void RunSchtasks(string arguments)
    {
        var psi = new System.Diagnostics.ProcessStartInfo("schtasks.exe", arguments)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        using var proc = System.Diagnostics.Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start schtasks.exe");

        var stdout = proc.StandardOutput.ReadToEnd();
        var stderr = proc.StandardError.ReadToEnd();
        proc.WaitForExit();

        if (proc.ExitCode != 0 && !stdout.Contains("does not exist", StringComparison.OrdinalIgnoreCase))
        {
            var parts = new[] { stdout.Trim(), stderr.Trim() }.Where(s => !string.IsNullOrEmpty(s));
            var msg = string.Join(" | ", parts);
            throw new InvalidOperationException($"schtasks.exe failed (exit {proc.ExitCode}): {msg}");
        }
    }
}
