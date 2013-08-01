//------------------------------------------------------------------------------
// <copyright file="DocumentCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Commands
{
    using System;
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The abstract class <see cref="DocumentCommand"/> can be extended to
    /// implement a concrete <i>Command</i> for a Microsoft Word document.
    /// </summary>
    public abstract class DocumentCommand : ICommand
    {
        /// <summary>
        /// The Microsoft Word document of this <i>Command</i>.
        /// </summary>
        protected readonly Word.Document Document;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentCommand"/>
        /// class with the specified <i>Receiver</i>.
        /// </summary>
        /// <param name="document">The <i>Receiver</i> of the <i>Command</i>.</param>
        /// <exception cref="ArgumentNullException">If the <c>application</c> argument is <c>null</c>.</exception>
        protected DocumentCommand(Word.Document document)
        {
            if (null == document)
            {
                throw new ArgumentNullException("document");
            }

            this.Document = document;
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public abstract void Execute();
    }
}
