//------------------------------------------------------------------------------
// <copyright file="UpdateFieldsCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using FlorianWolters.Office.Word.Commands;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="UpdateFieldsCommand"/> implements a <i>Command</i>
    /// which updates all <see cref="Word.Fields"/> of the active Microsoft Word
    /// document.
    /// </summary>
    internal class UpdateFieldsCommand : ApplicationCommand
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="UpdateFieldsCommand"/> class with the specified
        /// <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public UpdateFieldsCommand(Word.Application application)
            : base(application)
        {
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public override void Execute()
        {
            Word.Document document = this.Application.ActiveDocument;

            document.Fields.Update();
        }
    }
}
