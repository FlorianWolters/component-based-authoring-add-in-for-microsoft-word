//------------------------------------------------------------------------------
// <copyright file="DialogResults.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Dialogs
{
    /// <summary>
    /// The class <see cref="DialogResults"/> specifies identifiers to indicate
    /// the return value of a Microsoft Word application dialog box.
    /// </summary>
    public enum DialogResults
    {
        /// <summary>
        /// The dialog box return value is <b>Close</b>.
        /// </summary>
        Close = -2,

        /// <summary>
        /// The dialog box return value is <b>OK</b>.
        /// </summary>
        Ok = -1,

        /// <summary>
        /// The dialog box return value is <b>Cancel</b>.
        /// </summary>
        Cancel = 0,

        /// <summary>
        /// The dialog box return value is <b>FirstButton</b> (usually sent from
        /// the first button of the dialog).
        /// </summary>
        FirstButton = 1,

        /// <summary>
        /// The dialog box return value is <b>FirstButton</b> (usually sent from
        /// the second button of the dialog).
        /// </summary>
        SecondButton = 2,

        /// <summary>
        /// The dialog box return value is <b>FirstButton</b> (usually sent from
        /// the third button of the dialog).
        /// </summary>
        ThirdButton = 3
    }
}
