//------------------------------------------------------------------------------
// <copyright file="ActivateUpdateStylesOnOpenEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.EventHandlers
{
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Event.EventExceptionHandlers;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using Word = Microsoft.Office.Interop.Word;

    internal class ActivateUpdateStylesOnOpenEventHandler
        : CommandEventHandler,
        IDocumentBeforeCloseEventHandler,
        IDocumentOpenEventHandler
    {
        // TODO Duplicated source code with UpdateAttachedTemplateEventHandler.
        // But I don't know how to minimize code duplication further without
        // losing flexibility. It is possible to automatically implement the
        // "On[...]" methods in dependency of the implemented interfaces?

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="ActivateUpdateStylesOnOpenEventHandler"/> class.
        /// </summary>
        /// <param name="command">The <i>Command</i> to execute with this <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">The <i>Event Handler</i> used to handle exceptions.</param>
        public ActivateUpdateStylesOnOpenEventHandler(
            ICommand command, IEventExceptionHandler exceptionHandler)
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
