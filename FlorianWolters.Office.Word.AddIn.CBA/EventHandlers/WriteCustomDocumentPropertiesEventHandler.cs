//------------------------------------------------------------------------------
// <copyright file="WriteCustomDocumentPropertiesEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.EventHandlers
{
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="WriteCustomDocumentPropertiesEventHandler"/> implements <i>Event Handler</i> methods which
    /// execute the <see cref="WriteCustomDocumentPropertiesCommand"/>.
    /// </summary>
    internal class WriteCustomDocumentPropertiesEventHandler
        : CommandEventHandler, IDocumentBeforeSaveEventHandler, IDocumentOpenEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WriteCustomDocumentPropertiesEventHandler"/> class.
        /// </summary>
        /// <param name="command">The <i>Command</i> to execute with this <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">The <i>Event Handler</i> used to handle exceptions.</param>
        public WriteCustomDocumentPropertiesEventHandler(ICommand command, IExceptionHandler exceptionHandler)
            : base(command, exceptionHandler)
        {
        }

        /// <summary>
        /// Handles the event which occurs when a <see cref="Word.Document"/> is opened.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/>.</param>
        public void OnDocumentOpen(Word.Document document)
        {
            this.TryExecute();
        }

        /// <summary>
        /// Handles the event which occurs before a <see cref="Word.Document"/> is saved.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being saved.</param>
        /// <param name="saveAsUI"><c>true</c> if called via <i>Save As</i>; <c>false</c> if called via <i>Save</i>.</param>
        /// <param name="cancel"><c>true</c> to prevent the <see cref="Word.Document"/> from being saved; <c>false</c> otherwise.</param>
        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            // Solution taken from the following Microsoft Developer Network (MSDN) thread:
            // http://social.msdn.microsoft.com/Forums/4db4878e-a27c-4a6c-9c7d-984d918c0db5/how-to-call-some-code-after-document-is-saved-in-word

            // Check whether the document has changed since it was last saved.
            if (!document.Saved)
            {
                // Cancel this save operation.
                cancel = true;

                // Save the document and execute the Command.
                document.Save();
                this.TryExecute();
            }
        }
    }
}
