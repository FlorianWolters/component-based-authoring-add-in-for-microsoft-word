//------------------------------------------------------------------------------
// <copyright file="HWndWrapper.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// The class <see cref="HWndWrapper"/> wraps a <see cref="IntPtr"/>.
    /// <para>
    /// A <see cref="HWndWrapper"/> object can used to convert a <see
    /// cref="IntPtr"/> window handle to a <see cref="IWin32Window"/> window
    /// handle.
    /// </para>
    /// </summary>
    public class HWndWrapper : IWin32Window
    {
        /// <summary>
        /// The <see cref="IntPtr"/> wrapped by this <see cref="HWndWrapper"/>.
        /// </summary>
        private readonly IntPtr handle;

        /// <summary>
        /// Initializes a new instance of the <see cref="HWndWrapper"/> class
        /// which wraps the specified <see cref="IntPtr"/>.
        /// </summary>
        /// <param name="handle">The <see cref="IntPtr"/> to wrap.</param>
        public HWndWrapper(IntPtr handle)
        {
            this.handle = handle;
        }

        /// <summary>
        /// Gets the handle to the window.
        /// </summary>
        public IntPtr Handle
        {
            get
            {
                return this.handle;
            }
        }
    }
}
