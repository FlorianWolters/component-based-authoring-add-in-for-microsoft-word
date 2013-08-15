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
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="WriteCustomDocumentPropertiesCommand"/> implements a <i>Command</i> which sets custom
    /// document properties in the active Microsoft Word document.
    /// </summary>
    internal class WriteCustomDocumentPropertiesCommand : ApplicationCommand
    {
        /// <summary>
        /// The <see cref="Logger"/> of this class.
        /// </summary>
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The name of the custom document property which is used to store the directory path of the active document.
        /// </summary>
        private readonly string propertyNameForLastDirectoryPath;

        /// <summary>
        /// Initializes a new instance of the <see cref="WriteCustomDocumentPropertiesCommand"/> class with the
        /// specified <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public WriteCustomDocumentPropertiesCommand(Word.Application application) : base(application)
        {
            // TODO Fix violation of IoC
            this.propertyNameForLastDirectoryPath = Settings.Default.DocPropertyNameForLastDirectoryPath;
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
                    "The document \"" + document.FullName + "\" must be saved, to determine its directory path.");
            }

            string actualPropertyValue = string.Empty;

            // TODO Fix violation of IoC
            string expectedPropertyValue = new FieldFilePathTranslator().Encode(document.Path);
            CustomDocumentPropertyReader customDocumentPropertyReader = new CustomDocumentPropertyReader(document);

            if (customDocumentPropertyReader.Exists(this.propertyNameForLastDirectoryPath))
            {
                actualPropertyValue = customDocumentPropertyReader.Get<string>(this.propertyNameForLastDirectoryPath);
            }

            if (expectedPropertyValue != actualPropertyValue)
            {
                new CustomDocumentPropertyWriter(document)
                    .Set(this.propertyNameForLastDirectoryPath, expectedPropertyValue);
                this.logger.Info(
                    "Set the custom document property with the name \"" + this.propertyNameForLastDirectoryPath
                    + "\" to the value \"" + expectedPropertyValue + "\" in the document \"" + document.FullName + "\".");
            }
        }
    }
}
