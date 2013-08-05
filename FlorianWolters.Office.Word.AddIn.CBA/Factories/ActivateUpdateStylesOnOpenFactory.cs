﻿//------------------------------------------------------------------------------
// <copyright file="ActivateUpdateStylesOnOpenFactory.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Factories
{
    using FlorianWolters.Office.Word.AddIn.CBA.Commands;
    using FlorianWolters.Office.Word.AddIn.CBA.EventHandlers;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="ActivateUpdateStylesOnOpenFactory"/>
    /// implements a <i>FactoryMethod</i> to create instances of <see
    /// cref="UpdateAttachedTemplateEventHandler"/>.
    /// </summary>
    public sealed class ActivateUpdateStylesOnOpenFactory : EventHandlerFactory
    {
        /// <summary>
        /// The <i>Singleton</i> instance of the <see cref="ActivateUpdateStylesOnOpenFactory"/>.
        /// </summary>
        public static readonly ActivateUpdateStylesOnOpenFactory Instance = new ActivateUpdateStylesOnOpenFactory();

        /// <summary>
        /// Prevents a default instance of the <see cref="ActivateUpdateStylesOnOpenFactory"/> class from being created.
        /// </summary>
        private ActivateUpdateStylesOnOpenFactory()
        {
        }

        /// <summary>
        /// Creates the <i>Command</i> to inject into the <i>Event Handler</i>.
        /// </summary>
        /// <param name="application">The Microsoft Word application used by the <i>Command</i>.</param>
        /// <returns>The newly created <i>Command</i> instance.</returns>
        protected override ICommand CreateCommand(Word.Application application)
        {
            return new ActivateUpdateStylesOnOpenCommand(application);
        }

        /// <summary>
        /// Creates the <i>Event Handler</i> to return by this <i>Factory</i>.
        /// </summary>
        /// <param name="command">The <i>Command</i> to inject into the <i>Event Handler</i>.</param>
        /// <param name="exceptionHandler">The <i>Exception Handler</i> to use if an <see cref="Exception"/> inside an <i>Event Handler</i> occurs.</param>
        /// <returns>The newly created <i>Event Handler</i> instance.</returns>
        protected override IEventHandler CreateEventHandler(
            ICommand command,
            IExceptionHandler exceptionHandler)
        {
            return new ActivateUpdateStylesOnOpenEventHandler(
                command,
                exceptionHandler);
        }
    }
}