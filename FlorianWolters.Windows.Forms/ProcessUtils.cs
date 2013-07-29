//------------------------------------------------------------------------------
// <copyright file="ProcessUtils.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms
{
    using System;
    using System.Diagnostics;
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="ProcessUtils"/> contains utility methods to
    /// retrieve data from a <see cref="Process"/>.
    /// </summary>
    public static class ProcessUtils
    {
        /// <summary>
        /// Returns the window handle of the main window of the current <see
        /// cref="Process"/>.
        /// </summary>
        /// <returns>The system-generated window handle of the main window of the associated process.</returns>
        public static IWin32Window MainWindowWin32HandleOfCurrentProcess()
        {
            return MainWindowWin32HandleForProcess(Process.GetCurrentProcess());
        }

        /// <summary>
        /// Returns the window handle of the main window of the specified <see
        /// cref="Process"/>.
        /// </summary>
        /// <param name="process">The <see cref="Process"/> to check.</param>
        /// <returns>The system-generated window handle of the main window of the associated process.</returns>
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
