using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Item_And_Location_Data_V1
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        

        [DllImport("ole32.dll")]
        public static extern int CoMarshalInterThreadInterfaceInStream(
        [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] object pUnk,
        out IntPtr pStm);

        [DllImport("ole32.dll")]
        public static extern int CoGetInterfaceAndReleaseStream(
        IntPtr pStm,
        [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppv);


        [DllImport("kernel32.dll")]
        public static extern IntPtr GetCurrentThread();

        [DllImport("kernel32.dll")]
        public static extern uint GetThreadId(IntPtr thread);

        [DllImport("kernel32.dll")]
        public static extern uint GetProcessId(IntPtr process);

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        public static Process GetExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        [STAThread]
        static void Main()
        {
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            
            GC.Collect();
            Thread.Sleep(5000);
            GC.Collect();
        }
        
    }
}
