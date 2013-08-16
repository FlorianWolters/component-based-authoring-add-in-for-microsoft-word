//------------------------------------------------------------------------------
// <copyright file="UpdateFieldsEventHandler.cs" company="Florian Wolters">
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
    /// The class <see cref="UpdateFieldsEventHandler"/> implements <i>Event Handler</i> methods which execute the <see
    /// cref="UpdateFieldsCommand"/>.
    /// </summary>
    internal class UpdateFieldsEventHandler : CommandEventHandler, IDocumentOpenEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateFieldsEventHandler"/> class.
        /// </summary>
        /// <param name="command">The <i>Command</i> to execute with this <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">The <i>Event Handler</i> used to handle exceptions.</param>
        public UpdateFieldsEventHandler(ICommand command, IExceptionHandler exceptionHandler)
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
    }
}
