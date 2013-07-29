//------------------------------------------------------------------------------
// <copyright file="ICommand.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Commands
{
    /// <summary>
    /// The interface <see cref="ICommand"/> specifies the <i>Command</i> design
    /// searchPattern.
    /// </summary>
    public interface ICommand
    {
        /// <summary>
        /// Runs this <i>Command</i>.
        /// </summary>
        void Execute();
    }
}
