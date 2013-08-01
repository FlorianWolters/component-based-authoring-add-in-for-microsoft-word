//------------------------------------------------------------------------------
// <copyright file="UpdateAttachedTemplateEventHandler.cs" company="Florian Wolters">
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

    internal class UpdateAttachedTemplateEventHandler
        : CommandEventHandler,
        IDocumentBeforeCloseEventHandler,
        IDocumentOpenEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="UpdateAttachedTemplateEventHandler"/> class.
        /// </summary>
        /// <param name="command">The <i>Command</i> to execute with this <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">The <i>Event Handler</i> used to handle exceptions.</param>
        public UpdateAttachedTemplateEventHandler(
            ICommand command, IExceptionHandler exceptionHandler)
            : base(command, exceptionHandler)
        {
        }

        public void OnDocumentOpen(Word.Document document)
        {
            this.TryExecute();
        }

        public void OnDocumentBeforeClose(Word.Document document, ref bool target)
        {
            this.TryExecute();
        }
    }
}
