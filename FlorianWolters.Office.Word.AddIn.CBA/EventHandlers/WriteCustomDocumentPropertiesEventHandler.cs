//------------------------------------------------------------------------------
// <copyright file="WriteCustomDocumentPropertiesEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.EventHandlers
{
    using System;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.DocumentProperties;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using Word = Microsoft.Office.Interop.Word;

    internal class WriteCustomDocumentPropertiesEventHandler
        : CommandEventHandler, IDocumentBeforeSaveEventHandler, IDocumentOpenEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="WriteCustomDocumentPropertiesEventHandler"/> class.
        /// </summary>
        /// <param name="command">The <i>Command</i> to execute with this <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">The <i>Event Handler</i> used to handle exceptions.</param>
        public WriteCustomDocumentPropertiesEventHandler(
            ICommand command, IExceptionHandler exceptionHandler)
            : base(command, exceptionHandler)
        {
        }

        public void OnDocumentOpen(Word.Document document)
        {
            this.TryExecute();
        }

        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            // We must make sure that the document has already been saved, to
            // determine the directory path.
            if (document.IsSaved())
            {
                this.TryExecute();
            }
        }
    }
}
