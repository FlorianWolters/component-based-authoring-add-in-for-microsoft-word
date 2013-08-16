//------------------------------------------------------------------------------
// <copyright file="UpdateAttachedTemplateCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using System.IO;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Commands;
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The <i>Command</i> <see cref="UpdateAttachedTemplateCommand"/> sets the attached Microsoft Word template for
    /// the active Microsoft Word document.
    /// <para>
    /// The <i>Command</i> starts to search in the directory where the the active document is saved, and ends at the
    /// root of the drive, e.g. <c>D:</c>.
    /// </para>
    /// </summary>
    internal class UpdateAttachedTemplateCommand : ApplicationCommand
    {
        /// <summary>
        /// The <see cref="Logger"/> of this class.
        /// </summary>
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The file name (without a file extension) to search for.
        /// </summary>
        private readonly string fileNameWithoutExtension;

        /// <summary>
        /// The file extensions to search for.
        /// </summary>
        private readonly string[] fileExtensions;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateAttachedTemplateCommand"/> class with the specified
        /// <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public UpdateAttachedTemplateCommand(Word.Application application) : base(application)
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

            if (null == document || string.Empty == document.Path)
            {
                return;
            }

            try
            {
                string actualTemplateFileName = ((Word.Template)document.get_AttachedTemplate()).FullName;
                string templateFileName = this.RetrieveDocumentTemplateFilePath(document.Path);

                if (actualTemplateFileName != templateFileName)
                {
                    using (new StateCapture(document))
                    {
                        document.set_AttachedTemplate(templateFileName);
                    }

                    this.logger.Info(
                        "Set the attached template of the document \"" + document.FullName
                        + "\" to the template \"" + templateFileName + "\".");
                }
            }
            catch (FileNotFoundException)
            {
                throw new TemplateNotFoundException(
                    "Unable to locate a Word template with the file name \""
                    + this.fileNameWithoutExtension + ".[" + string.Join("|", this.fileExtensions)
                    + "]\" in the same directory as the Word document \""
                    + document.FullName + "\" (or in one of its parent directories).");
            }
        }

        /// <summary>
        /// Searches for the file path of the Word template file.
        /// </summary>
        /// <param name="directoryPath">The directory path to begin the search.</param>
        /// <returns>The absolute file path.</returns>
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
                    // We reached the root directory of the file system, e.g. "D:".
                    throw new FileNotFoundException();
                }

                result = this.RetrieveDocumentTemplateFilePath(directoryInfo.FullName);
            }

            return result;
        }
    }
}
