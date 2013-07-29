//------------------------------------------------------------------------------
// <copyright file="UpdateAttachedTemplateCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.ComponentAddIn.Commands
{
    using System.IO;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Commands;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The <i>Command</i> <see cref="UpdateAttachedTemplateCommand"/>
    /// automatically sets the document template file of the active document in
    /// a Microsoft Word application.
    /// <para>
    /// The <i>Command</i> starts to search in the directory where the the
    /// active document is saved, and ends at the root of the drive, e.g.
    /// <c>D:</c>.
    /// </para>
    /// </summary>
    internal class UpdateAttachedTemplateCommand : ApplicationCommand
    {
        private readonly string fileNameWithoutExtension;

        private readonly string[] fileExtensions;

        public UpdateAttachedTemplateCommand(Word.Application application)
            : base(application)
        {
            // TODO Fix violation of IoC.
            this.fileNameWithoutExtension = Settings.Default.WordTemplateFilename;
            this.fileExtensions = Settings.Default.WordTemplateFileExtensions.Split(';');
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public override void Execute()
        {
            Word.Document document = this.Application.ActiveDocument;

            if (string.Empty == document.Path)
            {
                return;
            }

            string templateFileName = this.RetrieveDocumentTemplateFilePath(document.Path);
            document.set_AttachedTemplate(templateFileName);
            document.UpdateStyles();
        }

        private string RetrieveDocumentTemplateFilePath(string directoryPath)
        {
            string result = string.Empty;
            bool found = false;

            foreach (string fileExtension in this.fileExtensions)
            {
                result = directoryPath + Path.DirectorySeparatorChar
                    + this.fileNameWithoutExtension + "." + fileExtension;

                if (File.Exists(result))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                DirectoryInfo directoryInfo = Directory.GetParent(directoryPath);
                if (null == directoryInfo)
                {
                    throw new FileNotFoundException(
                        "Unable to locate a Word Template file with the filename \""
                        + this.fileNameWithoutExtension + ".[" + string.Join("|", this.fileExtensions)
                        + "]\" in the directory of the Document (or one of its parent directories).");
                }

                result = this.RetrieveDocumentTemplateFilePath(directoryInfo.FullName);
            }

            return result;
        }
    }
}
