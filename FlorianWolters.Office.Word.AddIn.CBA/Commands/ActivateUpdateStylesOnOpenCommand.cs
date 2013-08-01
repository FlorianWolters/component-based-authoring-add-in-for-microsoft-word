//------------------------------------------------------------------------------
// <copyright file="ActivateUpdateStylesOnOpenCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using FlorianWolters.Office.Word.Commands;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The <i>Command</i> <see cref="ActivateUpdateStylesOnOpenCommand"/>
    /// activates the option for the active Microsoft Word document which
    /// determines if the styles of the active Microsoft Word document are
    /// updated to match the styles in the attached Microsoft Word template.
    /// </summary>
    internal class ActivateUpdateStylesOnOpenCommand : ApplicationCommand
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="ActivateUpdateStylesOnOpenCommand"/> class with the specified
        /// <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public ActivateUpdateStylesOnOpenCommand(Word.Application application)
            : base(application)
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
                document.UpdateStylesOnOpen = true;
            }
        }
    }
}
