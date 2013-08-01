//------------------------------------------------------------------------------
// <copyright file="ApplicationEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Event
{
    using System;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using Word = Microsoft.Office.Interop.Word;

    // TODO Find a better class name.

    /// <summary>
    /// The class <see cref="ApplicationEventHandler"/> allows to subscribe and
    /// unsubscribe <i>Events</i> do a Microsoft Office Word application.
    /// </summary>
    public class ApplicationEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="ApplicationEventHandler"/> class which allows to subscribe
        /// <i>Event Handlers</i> to the specified Microsoft Office Word
        /// application.
        /// </summary>
        /// <param name="application">The Microsoft Office Word application.</param>
        /// <exception cref="ArgumentNullException">If the <c>application</c> argument is <c>null</c>.</exception>
        public ApplicationEventHandler(Word.Application application)
        {
            if (null == application)
            {
                throw new ArgumentNullException("application");
            }

            this.Application = application;
        }

        /// <summary>
        /// Gets the Microsoft Office Word application to manage.
        /// </summary>
        public Word.Application Application { get; private set; }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to all <i>Events</i>
        /// it implements.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        public void SubscribeEventHandler(IEventHandler eventHandler)
        {
            // TODO I don't know how-to improve this further.
            // ATTENTION: This method MUST be modified if a new Word Application
            // event handler registration method is added to this class.
            if (eventHandler is IDocumentBeforeCloseEventHandler)
            {
                this.SubscribeEventHandler((IDocumentBeforeCloseEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentBeforePrintEventHandler)
            {
                this.SubscribeEventHandler((IDocumentBeforePrintEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentBeforeSaveEventHandler)
            {
                this.SubscribeEventHandler((IDocumentBeforeSaveEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentChangeEventHandler)
            {
                this.SubscribeEventHandler((IDocumentChangeEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentOpenEventHandler)
            {
                this.SubscribeEventHandler((IDocumentOpenEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentSyncEventHandler)
            {
                this.SubscribeEventHandler((IDocumentSyncEventHandler)eventHandler);
            }

            if (eventHandler is INewDocumentEventHandler)
            {
                this.SubscribeEventHandler((INewDocumentEventHandler)eventHandler);
            }

            if (eventHandler is IQuitEventHandler)
            {
                this.SubscribeEventHandler((IQuitEventHandler)eventHandler);
            }

            if (eventHandler is IStartupEventHandler)
            {
                this.SubscribeEventHandler((IStartupEventHandler)eventHandler);
            }

            if (eventHandler is IWindowActivateEventHandler)
            {
                this.SubscribeEventHandler((IWindowActivateEventHandler)eventHandler);
            }

            if (eventHandler is IWindowBeforeDoubleClickEventHandler)
            {
                this.SubscribeEventHandler((IWindowBeforeDoubleClickEventHandler)eventHandler);
            }

            if (eventHandler is IWindowBeforeRightClickEventHandler)
            {
                this.SubscribeEventHandler((IWindowBeforeRightClickEventHandler)eventHandler);
            }

            if (eventHandler is IWindowDeactivateEventHandler)
            {
                this.SubscribeEventHandler((IWindowDeactivateEventHandler)eventHandler);
            }

            if (eventHandler is IWindowSelectionChangeEventHandler)
            {
                this.SubscribeEventHandler((IWindowSelectionChangeEventHandler)eventHandler);
            }

            if (eventHandler is IWindowSizeEventHandler)
            {
                this.SubscribeEventHandler((IWindowSizeEventHandler)eventHandler);
            }

            if (eventHandler is IXMLSelectionChangeEventHandler)
            {
                this.SubscribeEventHandler((IXMLSelectionChangeEventHandler)eventHandler);
            }

            if (eventHandler is IXMLValidationErrorEventHandler)
            {
                this.SubscribeEventHandler((IXMLValidationErrorEventHandler)eventHandler);
            }
        }

        ////private void SubscribeIfTypeOf<T>(IEventHandler eventHandler)
        ////{
        ////    if (eventHandler.GetType() == typeof(T))
        ////    {
        ////        this.SubscribeEventHandler(typeof(T));
        ////    }
        ////}

        private void SubscribeEventHandler(IDocumentBeforeCloseEventHandler eventHandler)
        {
            this.Application.DocumentBeforeClose -= eventHandler.OnDocumentBeforeClose;
            this.Application.DocumentBeforeClose += eventHandler.OnDocumentBeforeClose;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> as an <i>Event
        /// Handler</i> which handles the <i>Event</i> that occurs before any
        /// open Document is printed.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IDocumentBeforePrintEventHandler eventHandler)
        {
            this.Application.DocumentBeforePrint -= eventHandler.OnDocumentBeforePrint;
            this.Application.DocumentBeforePrint += eventHandler.OnDocumentBeforePrint;
        }

        private void SubscribeEventHandler(IDocumentBeforeSaveEventHandler eventHandler)
        {
            this.Application.DocumentBeforeSave -= eventHandler.OnDocumentBeforeSave;
            this.Application.DocumentBeforeSave += eventHandler.OnDocumentBeforeSave;
        }

        private void SubscribeEventHandler(IDocumentChangeEventHandler eventHandler)
        {
            this.Application.DocumentChange -= eventHandler.OnDocumentChange;
            this.Application.DocumentChange += eventHandler.OnDocumentChange;
        }

        private void SubscribeEventHandler(IDocumentOpenEventHandler eventHandler)
        {
            this.Application.DocumentOpen -= eventHandler.OnDocumentOpen;
            this.Application.DocumentOpen += eventHandler.OnDocumentOpen;
        }

        private void SubscribeEventHandler(INewDocumentEventHandler eventHandler)
        {
            Word.ApplicationEvents4_Event wordEvent = this.Application;

            wordEvent.NewDocument -= eventHandler.OnNewDocument;
            wordEvent.NewDocument += eventHandler.OnNewDocument;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> as an <i>Event
        /// Handler</i> which handles the <i>Event</i> that occurs when the user
        /// quits Word. 
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IQuitEventHandler eventHandler)
        {
            Word.ApplicationEvents4_Event wordEvent = this.Application;

            wordEvent.Quit -= eventHandler.OnQuit;
            wordEvent.Quit += eventHandler.OnQuit;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> as an <i>Event
        /// Handler</i> which handles the <i>Event</i> that occurs when Word
        /// starts. 
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IStartupEventHandler eventHandler)
        {
            this.Application.Startup -= eventHandler.OnStartup;
            this.Application.Startup += eventHandler.OnStartup;
        }

        private void SubscribeEventHandler(IWindowActivateEventHandler eventHandler)
        {
            this.Application.WindowActivate -= eventHandler.OnWindowActivate;
            this.Application.WindowActivate += eventHandler.OnWindowActivate;
        }

        private void SubscribeEventHandler(IWindowDeactivateEventHandler eventHandler)
        {
            this.Application.WindowDeactivate -= eventHandler.OnWindowDeactivate;
            this.Application.WindowDeactivate += eventHandler.OnWindowDeactivate;
        }

        private void SubscribeEventHandler(IWindowSelectionChangeEventHandler eventHandler)
        {
            this.Application.WindowSelectionChange -= eventHandler.OnWindowSelectionChange;
            this.Application.WindowSelectionChange += eventHandler.OnWindowSelectionChange;
        }

        private void SubscribeEventHandler(IWindowSizeEventHandler eventHandler)
        {
            this.Application.WindowSize -= eventHandler.OnWindowSize;
            this.Application.WindowSize += eventHandler.OnWindowSize;
        }

        private void SubscribeEventHandler(IXMLSelectionChangeEventHandler eventHandler)
        {
            this.Application.XMLSelectionChange -= eventHandler.OnXMLSelectionChange;
            this.Application.XMLSelectionChange += eventHandler.OnXMLSelectionChange;
        }

        private void SubscribeEventHandler(IXMLValidationErrorEventHandler eventHandler)
        {
            this.Application.XMLValidationError -= eventHandler.OnXMLValidationError;
            this.Application.XMLValidationError += eventHandler.OnXMLValidationError;
        }
    }
}
