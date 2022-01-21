using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QC_Label
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dtpDate.Format = DateTimePickerFormat.Custom;
            dtpDate.CustomFormat = "dd-MM-yyyy";
        }

        class RawPrinterHelper
        {
            // Structure and API declarions:
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public class DOCINFOA
            {
                [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
                [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
                [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
            }
            [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);

            [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool ClosePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

            [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndDocPrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartPagePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndPagePrinter(IntPtr hPrinter);

            [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);

            // SendBytesToPrinter()
            // When the function is given a printer name and an unmanaged array
            // of bytes, the function sends those bytes to the print queue.
            // Returns true on success, false on failure.
            public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
            {
                Int32 dwError = 0, dwWritten = 0;
                IntPtr hPrinter = new IntPtr(0);
                DOCINFOA di = new DOCINFOA();
                bool bSuccess = false; // Assume failure unless you specifically succeed.

                di.pDocName = "My C#.NET RAW Document";
                di.pDataType = "RAW";

                // Open the printer.
                if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
                {
                    // Start a document.
                    if (StartDocPrinter(hPrinter, 1, di))
                    {
                        // Start a page.
                        if (StartPagePrinter(hPrinter))
                        {
                            // Write your bytes.
                            bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                            EndPagePrinter(hPrinter);
                        }
                        EndDocPrinter(hPrinter);
                    }
                    ClosePrinter(hPrinter);
                }
                // If you did not succeed, GetLastError may give more information
                // about why not.
                if (bSuccess == false)
                {
                    dwError = Marshal.GetLastWin32Error();
                }
                return bSuccess;
            }

            public static bool SendFileToPrinter(string szPrinterName, string szFileName)
            {
                // Open the file.
                FileStream fs = new FileStream(szFileName, FileMode.Open);
                // Create a BinaryReader on the file.
                BinaryReader br = new BinaryReader(fs);
                // Dim an array of bytes big enough to hold the file's contents.
                Byte[] bytes = new Byte[fs.Length];
                bool bSuccess = false;
                // Your unmanaged pointer.
                IntPtr pUnmanagedBytes = new IntPtr(0);
                int nLength;

                nLength = Convert.ToInt32(fs.Length);
                // Read the contents of the file into the array.
                bytes = br.ReadBytes(nLength);
                // Allocate some unmanaged memory for those bytes.
                pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
                // Copy the managed byte array into the unmanaged array.
                Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
                // Send the unmanaged bytes to the printer.
                bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
                // Free the unmanaged memory that you allocated earlier.
                Marshal.FreeCoTaskMem(pUnmanagedBytes);
                return bSuccess;
            }
            public static bool SendStringToPrinter(string szPrinterName, string szString)
            {
                IntPtr pBytes;
                Int32 dwCount;
                // How many characters are in the string?
                dwCount = szString.Length;
                // Assume that the printer is expecting ANSI text, and then convert
                // the string to ANSI text.
                pBytes = Marshal.StringToCoTaskMemAnsi(szString);
                // Send the converted ANSI string to the printer.
                SendBytesToPrinter(szPrinterName, pBytes, dwCount);
                Marshal.FreeCoTaskMem(pBytes);
                return true;
            }

            public static bool SendTextFileToPrinter(string szFileName, string printerName)
            {
                var sb = new StringBuilder();

                using (var sr = new StreamReader(szFileName, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                        sb.AppendLine(sr.ReadLine());
                    }
                }

                return RawPrinterHelper.SendStringToPrinter(printerName, sb.ToString());
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string strQRCode = "";
            string strDate = dtpDate.Value.ToString("yyMMdd");
            string strQuantity = tbQuantity.Text.Trim();
            while (strQuantity.Length < 5)
            {
                strQuantity = "0" + strQuantity;
            }
            string strCavity=tbCavity.Text.ToUpper().Trim();
            while (strCavity.Length < 3)
            {
                strCavity="0" + strCavity;
            }
            strQRCode = tbSup.Text.ToUpper().Trim() + "000" + tbMaterial.Text.ToUpper().Trim() + strDate + tbSerial.Text.ToUpper() + strQuantity + strCavity;
            string prin = "";
            prin = "CT~~CD,~CC^~CT~" +
"^XA" +
"~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4^MD30^LRN^CI0" +
"^MMT" +
"^PW599" +
"^LL0240" +
"^LS0" +
"^FO0,0^GFA,01280,01280,00020,:Z64:" +
"eJzt0sFKw0AQANAJC82lJNeCa/sL6U1oMB/TH0gJ2EKpHU/evBeaf6gfIBY8eLM/kMKKBy/aLvRgpDFj1t3EiGfx0rksPJiZnWEADvGfEZqXYc1GAIKhYy0hIGks1AaiMlvZlmETlnBiCtmhBH/bRx5QzQTwTYQtz/62nLL2Zo/dRc2IMifvYVfaVQ93sE6c7Aj7t6TMVr90onUCLwxXXkPlahvGifpf7ulS2m6SIN/hg2dhaTDsFbbHZkD3RKRnOxsnlhhjazK7rOw8mluih95khm65l/xtziRbhDVz3585y5gkZZRq2624nQepLKxjzBlecZcIYUrIy6VHbd5+3SN8zNA3c0Do89NiDnetTBobcV+wi7QwqCz11byPsTKhbbBLlVnxNYIlTI+nL2P8GKFjbHAnlTV4kdZZ/riDJrDfN3KIv4tPGOGjig==:E7F3" +
"^FO0,32^GFA,01280,01280,00020,:Z64:" +
"eJzt0jFOwzAUBuAXecjmHACTXIEVNWquAuICiSLRoFb0cQG4QLkDK0OlSAxs9AIustSBBdpUHZqqqY1b24WhK2LpWyx9suX/PT2AY/1npYesABAEQR/Bg7PUWHVG1MiQv7WZNlUTSZ1VEM+u0FMNsNCZADbN0SsLYJEyJlUTTtdIygpYYk2phsoW+nf6/nm5syAbc9qcYDARRK1KH7XRfMzhk2C0LImaW+s88m2+Xh8hvnQ25IlcYJEAxJ4x6LS0rRHa2qLGBLzuck90EWqA/ik3dpsPPNHSfUjoD9+NyeWAVOTJ16Y2bybf6oORhlSRoN7m2dpixHyZ1Epbc2OGRTv3LFAKe4ISntkB5iELv9Z4oW0wfzX50pi1t3MWFNjeChY7s71BWsfgzPWWLeqd6Xza7B+TH/OtZS/V/q2zX3vADqzEsf6ovgHKMada:7317" +
"^FO0,64^GFA,01280,01280,00020,:Z64:" +
"eJzt0TFKxUAQBuA/bJFGkvYJ4h7CJmB8exXBCwRsAoaXBcHWCwhewdJCMPCKXGPEIp1ZsIk8yTgxJlkvoM0b2JD9mGFnZ4F9/Hskv7epLEqgLMGwm1MOxK7uSH6z0XJJFTt98qxkOmZ25tGr5dFOHrw8AkVSy8zVZAENOtpUqyqQEgtf6mEb2uHDjmKx+L1WkvttunRkxPRusQgZpS1T+bzUQgxinW68C4u9WlzctkvPgymL8KZFkvz0LOcauQf3b97dJuuXPCoJukKUbDyLCWdbRHmx1JKsj8bqLB3mO/ZHWCnHMFT4dqRdgcOul/F2s5lsDbVrrWdpesk2vq+r+UmZPwNu3MZcb2Uu0xs51QcuXweLjfM/xwr7+Mv4AjOFsf8=:217B" +
"^FT489,160^A0N,102,100^FB98,1,0^FH\\^FD01^FS" +
"^FT169,44^A0N,28,28^FB168,1,0^FH\\^FD"+tbMaterial.Text.ToUpper().Trim()+"^FS" +
"^FT169,76^A0N,28,28^FB103,1,0^FH\\^FD"+tbMate.Text.ToUpper().Trim()+"^FS" +
"^FT169,108^A0N,28,28^FB42,1,0^FH\\^FD"+tbQuantity.Text+"^FS" +
"^FT169,155^A0N,28,26^FB278,1,0^FH\\^FD"+ tbSup.Text.ToUpper().Trim() + "000" + tbMaterial.Text.ToUpper().Trim() + "^FS" +
"^FT31,242^BQN,2,4" +
"^FH\\^FDLA,"+strQRCode+"^FS" +
"^FT169,189^A0N,28,26^FB236,1,0^FH\\^FD"+ strDate + tbSerial.Text.ToUpper() + strQuantity + strCavity + "^FS" +
"^XZ";
            PrintDialog pd = new PrintDialog();
            pd.Document = new PrintDocument();
            pd.PrinterSettings = new PrinterSettings();

            RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, prin);
        }
    }
}
