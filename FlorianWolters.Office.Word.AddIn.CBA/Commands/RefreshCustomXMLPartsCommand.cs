//------------------------------------------------------------------------------
// <copyright file="RefreshCustomXMLPartsCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using System.IO;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.CustomXML;
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The <i>Command</i> <see cref="RefreshCustomXMLPartsCommand"/> refreshes all custom XML parts in the active
    /// Microsoft Word document with the content from the XML files in a subdirectory of that document.
    /// </summary>
    internal class RefreshCustomXMLPartsCommand : ApplicationCommand
    {
        /// <summary>
        /// The <see cref="Logger"/> of this class.
        /// </summary>
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The name of the directory with the XML files.
        /// </summary>
        private readonly string xmlDirectoryName;

        /// <summary>
        /// Initializes a new instance of the <see cref="RefreshCustomXMLPartsCommand"/> class with the specified
        /// <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public RefreshCustomXMLPartsCommand(Word.Application application) : base(application)
        {
            // TODO Fix violation of IoC.
            this.xmlDirectoryName = Settings.Default.XMLDirectoryName;
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public override void Execute()
        {
            Word.Document document = this.Application.ActiveDocument;

            if (null == document || string.Empty == document.Path)
            {
                return;
            }

            string directoryPath = this.DirectoryPathForXMLFiles(document);
            
            if (Directory.Exists(directoryPath))
            {
                // TODO Fix violation of IoC.
                CustomXMLPartRepository customXMLPartRepository = new CustomXMLPartRepository(document.CustomXMLParts);
                customXMLPartRepository.SynchronizeWithDirectory(directoryPath);

                this.logger.Info(
                    "Synchronized the custom XML parts of the document \""
                    + this.Application.ActiveDocument.FullName + "\" with the directory \"" + directoryPath + "\".");
            }
        }

        /// <summary>
        /// Retrieves the absolute directory path of the directory with the XML files for the specified <see
        /// cref="Word.Document"/>.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> whose XML directory path to retrieve.</param>
        /// <returns>
        /// The absolute directory path of the XML directory for the specified <see cref="Word.Document"/>.
        /// </returns>
        private string DirectoryPathForXMLFiles(Word.Document document)
        {
            return document.Path + Path.DirectorySeparatorChar + this.xmlDirectoryName;
        }
    }
}
