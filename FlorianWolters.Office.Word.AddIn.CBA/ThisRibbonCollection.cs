//------------------------------------------------------------------------------
// <copyright file="ThisRibbonCollection.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA
{
    /// <summary>
    /// The class <see cref="ThisRibbonCollection"/> contains all ribbons of the
    /// application.
    /// </summary>
    internal partial class ThisRibbonCollection
    {
        /// <summary>
        /// Gets the ribbon.
        /// </summary>
        internal Ribbon ThisRibbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
