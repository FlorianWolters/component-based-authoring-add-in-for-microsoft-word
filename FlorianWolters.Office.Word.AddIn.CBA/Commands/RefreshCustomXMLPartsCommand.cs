//------------------------------------------------------------------------------
// <copyright file="RefreshCustomXMLPartsCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using System.IO;
    using FlorianWolters.Office.Word.AddIn.CBA.CustomXML;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Commands;
    using Word = Microsoft.Office.Interop.Word;

    internal class RefreshCustomXMLPartsCommand : ApplicationCommand
    {
        private readonly string xmlDirectoryName;

        public RefreshCustomXMLPartsCommand(Word.Application application)
            : base(application)
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

            CustomXMLPartRepository customXMLPartRepository = new CustomXMLPartRepository(document.CustomXMLParts);

            // The approach is very simple:
            // 1. Delete all custom XML parts which are user defined.
            // 2. Add all XML files in the specified subdirectory as new custom XML parts.
            customXMLPartRepository.DeleteAllNotBuiltIn();
            string directoryPath = this.DirectoryPathForXMLFiles(document);
            customXMLPartRepository.AddFromDirectory(directoryPath);
        }

        private string DirectoryPathForXMLFiles(Word.Document document)
        {
            return document.Path + Path.DirectorySeparatorChar + this.xmlDirectoryName;
        }
    }
}
