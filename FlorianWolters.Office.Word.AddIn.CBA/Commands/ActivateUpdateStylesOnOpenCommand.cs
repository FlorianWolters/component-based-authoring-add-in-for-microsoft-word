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

            if (string.Empty != document.Path && !document.UpdateStylesOnOpen)
            {
                document.UpdateStylesOnOpen = true;
            }
        }
    }
}
