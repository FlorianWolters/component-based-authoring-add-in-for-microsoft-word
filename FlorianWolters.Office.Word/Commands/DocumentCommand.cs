//------------------------------------------------------------------------------
// <copyright file="DocumentCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Commands
{
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The abstract class <see cref="DocumentCommand"/> can be extended to
    /// implement a concrete <i>Command</i> for a Microsoft Word document.
    /// </summary>
    public abstract class DocumentCommand : ICommand
    {
        /// <summary>
        /// The logger for the class <see cref="DocumentCommand"/>.
        /// </summary>
        protected static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The Microsoft Word Document of this Command.
        /// </summary>
        protected readonly Word.Document Document;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentCommand"/>
        /// class with the specified <i>Receiver</i>.
        /// </summary>
        /// <param name="document">The <i>Receiver</i> of the <i>Command</i>.</param>
        protected DocumentCommand(Word.Document document)
        {
            this.Document = document;
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public abstract void Execute();
    }
}
