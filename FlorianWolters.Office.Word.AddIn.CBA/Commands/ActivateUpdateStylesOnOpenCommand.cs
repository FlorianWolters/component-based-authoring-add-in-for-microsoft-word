//------------------------------------------------------------------------------
// <copyright file="ActivateUpdateStylesOnOpenCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using FlorianWolters.Office.Word.Commands;
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The <i>Command</i> <see cref="ActivateUpdateStylesOnOpenCommand"/> copies all styles from the attached Microsoft
    /// Word template into the active Microsoft Word document, overwriting any existing styles in the document that have
    /// the same name. 
    /// <para>
    /// In addition, this <i>Command</i> activates the option of the active Microsoft Word document which determines if
    /// the styles of the document are updated to match the styles in the attached Microsoft Word template, each time
    /// the document is opened.
    /// </para>
    /// </summary>
    internal class ActivateUpdateStylesOnOpenCommand : ApplicationCommand
    {
        /// <summary>
        /// The <see cref="Logger"/> of this class.
        /// </summary>
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Initializes a new instance of the <see cref="ActivateUpdateStylesOnOpenCommand"/> class with the specified
        /// <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public ActivateUpdateStylesOnOpenCommand(Word.Application application) : base(application)
        {
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public override void Execute()
        {
            Word.Document document = this.Application.ActiveDocument;

            if (null != document && !document.UpdateStylesOnOpen)
            {
                using (new StateCapture(document))
                {
                    document.UpdateStylesOnOpen = true;
                    document.UpdateStyles();
                }

                this.logger.Info(
                    "The styles of the document \"" + document.FullName
                    + "\" have been updated to match the styles in the attached template \""
                    + ((Word.Template)document.get_AttachedTemplate()).FullName + "\".");
            }
        }
    }
}
