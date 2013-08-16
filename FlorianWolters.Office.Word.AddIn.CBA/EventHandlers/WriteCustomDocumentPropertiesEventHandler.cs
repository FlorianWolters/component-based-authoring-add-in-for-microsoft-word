//------------------------------------------------------------------------------
// <copyright file="WriteCustomDocumentPropertiesEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.EventHandlers
{
    using System;
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
        /// <param name="document">The <see cref="Word.Document"/> that's being opened.</param>
        public void OnDocumentOpen(Word.Document document)
        {
            this.TryExecute();
        }

        /// <summary>
        /// Handles the event which occurs before any <see cref="Word.Document"/> is saved.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being saved.</param>
        /// <param name="saveAsUI">
        /// <c>true</c> if the <b>Save As</b> dialog box is displayed, whether to save a new <see
        /// cref="Word.Document"/>, in response to the <b>Save</b> command; or in response to the <b>Save As</b>
        /// command; or in response to the <b>SaveAs</b> or <b>SaveAs2</b> method.
        /// </param>
        /// <param name="cancel">
        /// <c>false</c> when the event occurs. If the event procedure sets this argument to <c>true</c>, the <see
        /// cref="Word.Document"/> is not saved when the procedure is finished.
        /// </param>
        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            // Solution taken from the following Microsoft Developer Network (MSDN) thread:
            // http://social.msdn.microsoft.com/Forums/en-US/4db4878e-a27c-4a6c-9c7d-984d918c0db5/how-to-call-some-code-after-document-is-saved-in-word
            try
            {
                // Check whether the document has changed since it was last saved.
                if (!document.Saved)
                {
                    // Cancel this save operation.
                    cancel = true;

                    // Save the document and execute the Command.
                    document.Save();
                    this.Command.Execute();
                }
            }
            catch (Exception ex)
            {
                // TODO This solution isn't perfect. If the method Save is called, a possible save failuere is not 
                // handled by the Microsoft Word application. Therefore this differs from the default behaviour.
                this.ExceptionHandler.HandleException(ex);
            }
        }
    }
}
