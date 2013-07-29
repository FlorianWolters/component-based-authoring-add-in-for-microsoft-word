//------------------------------------------------------------------------------
// <copyright file="ApplicationCommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Commands
{
    using NLog;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The abstract class <see cref="ApplicationCommand"/> can be extended to
    /// implement a concrete <i>Command</i> for a Microsoft Word application.
    /// </summary>
    public abstract class ApplicationCommand : ICommand
    {
        /// <summary>
        /// The logger for the class <see cref="ApplicationCommand"/>.
        /// </summary>
        protected static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The Microsoft Word Application of this Command.
        /// </summary>
        protected readonly Word.Application Application;

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="ApplicationCommand"/> class with the specified <i>Receiver</i>.
        /// </summary>
        /// <param name="application">The <i>Receiver</i> of the <i>Command</i>.</param>
        protected ApplicationCommand(Word.Application application)
        {
            this.Application = application;
        }

        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        public abstract void Execute();
    }
}
