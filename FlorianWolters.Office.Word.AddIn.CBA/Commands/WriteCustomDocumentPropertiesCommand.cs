//------------------------------------------------------------------------------
// <copyright file="WriteCustomDocumentPropertiesCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using System;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.DocumentProperties;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="WriteCustomDocumentPropertiesCommand"/> implements
    /// a <i>Command</i> which sets custom document properties in the active
    /// Microsoft Word document.
    /// </summary>
    internal class WriteCustomDocumentPropertiesCommand : ApplicationCommand
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="WriteCustomDocumentPropertiesCommand"/> class with the
        /// specified <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public WriteCustomDocumentPropertiesCommand(Word.Application application)
            : base(application)
        {
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public override void Execute()
        {
            Word.Document document = this.Application.ActiveDocument;

            if (!this.Application.ActiveDocument.IsSaved())
            {
                throw new Exception(
                    "The document must be saved, to determine its directory path.");
            }

            // TODO Fix violation of IoC
            string propertyName = Settings.Default.DocPropertyNameForLastDirectoryPath;
            string propertyValue = new FieldFilePathTranslator()
                .Encode(document.Path);
            new CustomDocumentPropertyWriter(document)
                .Set(propertyName, propertyValue);
        }
    }
}
