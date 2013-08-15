//------------------------------------------------------------------------------
// <copyright file="UpdateFieldsCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Commands
{
    using System;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Fields;
    using FlorianWolters.Office.Word.Fields.UpdateStrategies;
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="UpdateFieldsCommand"/> implements a <i>Command</i> which updates all <see
    /// cref="Word.Fields"/> of the active Microsoft Word document.
    /// </summary>
    internal class UpdateFieldsCommand : ApplicationCommand
    {
        /// <summary>
        /// The <see cref="Logger"/> of this class.
        /// </summary>
        private readonly Logger logger = LogManager.GetCurrentClassLogger();

        private readonly FieldUpdater fieldUpdater;

        /// <summary>
        /// Initializes a new instance of the <see cref="UpdateFieldsCommand"/> class with the specified
        /// <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        public UpdateFieldsCommand(Word.Application application)
            : base(application)
        {
            this.fieldUpdater = new FieldUpdater(new UpdateTarget());
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public override void Execute()
        {
            Word.Document document = this.Application.ActiveDocument;
            
            if (null == document)
            {
                throw new InvalidOperationException("The Microsoft Word application has no active document.");
            }

            using (new StateCapture(document))
            {
                this.fieldUpdater.Update(document);
                this.logger.Info("Updated the result of all fields in the document \"" + document.FullName + "\".");
            }
        }
    }
}
