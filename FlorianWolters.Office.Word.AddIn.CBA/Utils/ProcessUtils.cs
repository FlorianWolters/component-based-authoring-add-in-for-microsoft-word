//------------------------------------------------------------------------------
// <copyright file="ProcessUtils.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Utils
{
    using System;
    using System.Diagnostics;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Forms;

    public static class ProcessUtils
    {
        public static IWin32Window MainWindowWin32HandleOfCurrentProcess()
        {
            return MainWindowWin32HandleForProcess(Process.GetCurrentProcess());
        }

        public static IWin32Window MainWindowWin32HandleForProcess(Process process)
        {
            IntPtr hwnd = process.MainWindowHandle;
            IWin32Window result = new HWndWrapper(hwnd);

            // The following does cause a System.AppDomainUnloadedException
            // exception if the Add-in is closed.
            // (http://social.msdn.microsoft.com/Forums/office/en-US/2859d53c-3e40-487d-b39e-e7a4e2dd75a6/word-2010-addin-shutdown-dispatcher-and-garbage-collection)
            // NativeWindow result = new NativeWindow();
            // result.AssignHandle(hwnd);
            return result;
        }
    }
}
