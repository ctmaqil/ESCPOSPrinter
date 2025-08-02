using System;
using System.Runtime.InteropServices;

namespace ESCPrintApp
{
    public class RawPrinterHelper
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
        }

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi)]
        public static extern bool OpenPrinter(string szPrinter, out IntPtr hPrinter, IntPtr pd);

        [DllImport("winspool.Drv", SetLastError = true)]
        public static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi)]
        public static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

        [DllImport("winspool.Drv", SetLastError = true)]
        public static extern bool EndDocPrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", SetLastError = true)]
        public static extern bool StartPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", SetLastError = true)]
        public static extern bool EndPagePrinter(IntPtr hPrinter);

        [DllImport("winspool.Drv", SetLastError = true)]
        public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

        public static bool SendBytesToPrinter(string szPrinterName, byte[] pBytes)
        {
            IntPtr pUnmanagedBytes = IntPtr.Zero;
            IntPtr hPrinter = IntPtr.Zero;
            int dwWritten = 0;
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false;

            di.pDocName = "ESC/POS Print Job";
            di.pDataType = "RAW";

            try
            {
                if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
                {
                    if (StartDocPrinter(hPrinter, 1, di))
                    {
                        if (StartPagePrinter(hPrinter))
                        {
                            pUnmanagedBytes = Marshal.AllocCoTaskMem(pBytes.Length);
                            Marshal.Copy(pBytes, 0, pUnmanagedBytes, pBytes.Length);
                            bSuccess = WritePrinter(hPrinter, pUnmanagedBytes, pBytes.Length, out dwWritten);
                            EndPagePrinter(hPrinter);
                        }
                        EndDocPrinter(hPrinter);
                    }
                    ClosePrinter(hPrinter);
                }

                if (!bSuccess)
                {
                    int dwError = Marshal.GetLastWin32Error();
                    throw new Exception($"Failed to send data to printer. Error code: {dwError}");
                }

                return bSuccess;
            }
            finally
            {
                if (pUnmanagedBytes != IntPtr.Zero)
                    try { Marshal.FreeCoTaskMem(pUnmanagedBytes); } catch { }
                if (hPrinter != IntPtr.Zero)
                    try { ClosePrinter(hPrinter); } catch { }
            }
        }
    }
}
