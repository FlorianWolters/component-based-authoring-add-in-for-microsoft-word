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
    using FlorianWolters.Office.Word.DocumentProperties;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using Word = Microsoft.Office.Interop.Word;

    internal class WriteCustomDocumentPropertiesEventHandler
        : IEventHandler, IDocumentBeforeSaveEventHandler, IDocumentOpenEventHandler
    {
        private readonly Word.Application application;

        public WriteCustomDocumentPropertiesEventHandler(Word.Application application)
        {
            this.application = application;
        }

        public void OnDocumentOpen(Word.Document document)
        {
            this.SetLastDirectoryPathCustomDocumentProperty(document);
        }

        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            // We must make sure that the document has already been saved, to
            // determine the directory path.
            if (document.IsSaved())
            {
                this.SetLastDirectoryPathCustomDocumentProperty(document);
            }
        }

        private void SetLastDirectoryPathCustomDocumentProperty(Word.Document document)
        {
            if (!document.IsSaved())
            {
                throw new Exception("The Document must be saved, to determine its directory path.");
            }

            string propertyName = Settings.Default.DocPropertyNameForLastDirectoryPath;

            // TODO Fix violation of IoC
            string propertyValue = new FieldFilePathTranslator().Encode(document.Path);

            new CustomDocumentPropertyWriter(document)
                .Set(propertyName, propertyValue);
        }
    }
}
