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

    /// <summary>
    /// The class <see cref="ApplicationEventHandler"/> allows to subscribe and unsubscribe <i>Event Handler</i> methods
    /// for a Microsoft Office Word application.
    /// </summary>
    public class ApplicationEventHandler
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ApplicationEventHandler"/> class which allows to subscribe
        /// <i>Event Handler</i> methods for the specified Microsoft Office Word application.
        /// </summary>
        /// <param name="application">The Microsoft Office Word application whose events to handle.</param>
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

        // TODO Find a way to implement the logic of this class with less source code.
        // TODO The two methods SubscribeEventHandler(IEventHandler eventHandler) and UnsubscribeEventHandler(IEventHandler eventHandler)
        // must be modified if a new event handler registration method is added.

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to all <i>Events</i> it implements.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        public void SubscribeEventHandler(IEventHandler eventHandler)
        {
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

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from all <i>Events</i> it implements.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        public void UnsubscribeEventHandler(IEventHandler eventHandler)
        {
            if (eventHandler is IDocumentBeforeCloseEventHandler)
            {
                this.UnsubscribeEventHandler((IDocumentBeforeCloseEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentBeforePrintEventHandler)
            {
                this.UnsubscribeEventHandler((IDocumentBeforePrintEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentBeforeSaveEventHandler)
            {
                this.UnsubscribeEventHandler((IDocumentBeforeSaveEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentChangeEventHandler)
            {
                this.UnsubscribeEventHandler((IDocumentChangeEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentOpenEventHandler)
            {
                this.UnsubscribeEventHandler((IDocumentOpenEventHandler)eventHandler);
            }

            if (eventHandler is IDocumentSyncEventHandler)
            {
                this.UnsubscribeEventHandler((IDocumentSyncEventHandler)eventHandler);
            }

            if (eventHandler is INewDocumentEventHandler)
            {
                this.UnsubscribeEventHandler((INewDocumentEventHandler)eventHandler);
            }

            if (eventHandler is IQuitEventHandler)
            {
                this.UnsubscribeEventHandler((IQuitEventHandler)eventHandler);
            }

            if (eventHandler is IStartupEventHandler)
            {
                this.UnsubscribeEventHandler((IStartupEventHandler)eventHandler);
            }

            if (eventHandler is IWindowActivateEventHandler)
            {
                this.UnsubscribeEventHandler((IWindowActivateEventHandler)eventHandler);
            }

            if (eventHandler is IWindowBeforeDoubleClickEventHandler)
            {
                this.UnsubscribeEventHandler((IWindowBeforeDoubleClickEventHandler)eventHandler);
            }

            if (eventHandler is IWindowBeforeRightClickEventHandler)
            {
                this.UnsubscribeEventHandler((IWindowBeforeRightClickEventHandler)eventHandler);
            }

            if (eventHandler is IWindowDeactivateEventHandler)
            {
                this.UnsubscribeEventHandler((IWindowDeactivateEventHandler)eventHandler);
            }

            if (eventHandler is IWindowSelectionChangeEventHandler)
            {
                this.UnsubscribeEventHandler((IWindowSelectionChangeEventHandler)eventHandler);
            }

            if (eventHandler is IWindowSizeEventHandler)
            {
                this.UnsubscribeEventHandler((IWindowSizeEventHandler)eventHandler);
            }

            if (eventHandler is IXMLSelectionChangeEventHandler)
            {
                this.UnsubscribeEventHandler((IXMLSelectionChangeEventHandler)eventHandler);
            }

            if (eventHandler is IXMLValidationErrorEventHandler)
            {
                this.UnsubscribeEventHandler((IXMLValidationErrorEventHandler)eventHandler);
            }
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>DocumentBeforeClose</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IDocumentBeforeCloseEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.DocumentBeforeClose += eventHandler.OnDocumentBeforeClose;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>DocumentBeforeClose</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IDocumentBeforeCloseEventHandler eventHandler)
        {
            this.Application.DocumentBeforeClose -= eventHandler.OnDocumentBeforeClose;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>DocumentBeforePrint</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IDocumentBeforePrintEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.DocumentBeforePrint += eventHandler.OnDocumentBeforePrint;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>DocumentBeforePrint</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IDocumentBeforePrintEventHandler eventHandler)
        {
            this.Application.DocumentBeforePrint -= eventHandler.OnDocumentBeforePrint;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>DocumentBeforeSave</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IDocumentBeforeSaveEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.DocumentBeforeSave += eventHandler.OnDocumentBeforeSave;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>DocumentBeforeSave</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IDocumentBeforeSaveEventHandler eventHandler)
        {
            this.Application.DocumentBeforeSave -= eventHandler.OnDocumentBeforeSave;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>DocumentChange</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IDocumentChangeEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.DocumentChange += eventHandler.OnDocumentChange;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>DocumentChange</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IDocumentChangeEventHandler eventHandler)
        {
            this.Application.DocumentChange -= eventHandler.OnDocumentChange;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>DocumentOpen</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IDocumentOpenEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.DocumentOpen += eventHandler.OnDocumentOpen;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>DocumentOpen</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IDocumentOpenEventHandler eventHandler)
        {
            this.Application.DocumentOpen -= eventHandler.OnDocumentOpen;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>NewDocument</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(INewDocumentEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += eventHandler.OnNewDocument;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>NewDocument</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(INewDocumentEventHandler eventHandler)
        {
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument -= eventHandler.OnNewDocument;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>Quit</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IQuitEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            ((Word.ApplicationEvents4_Event)this.Application).Quit += eventHandler.OnQuit;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>Quit</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IQuitEventHandler eventHandler)
        {
            ((Word.ApplicationEvents4_Event)this.Application).Quit -= eventHandler.OnQuit;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>Startup</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IStartupEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.Startup += eventHandler.OnStartup;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>Startup</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IStartupEventHandler eventHandler)
        {
            this.Application.Startup -= eventHandler.OnStartup;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>WindowActivate</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IWindowActivateEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.WindowActivate += eventHandler.OnWindowActivate;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>WindowActivate</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IWindowActivateEventHandler eventHandler)
        {
            this.Application.WindowActivate -= eventHandler.OnWindowActivate;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>WindowDeactivate</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IWindowDeactivateEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.WindowDeactivate += eventHandler.OnWindowDeactivate;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>WindowDeactivate</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IWindowDeactivateEventHandler eventHandler)
        {
            this.Application.WindowDeactivate -= eventHandler.OnWindowDeactivate;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>WindowSelectionChange</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IWindowSelectionChangeEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.WindowSelectionChange += eventHandler.OnWindowSelectionChange;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>WindowSelectionChange</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IWindowSelectionChangeEventHandler eventHandler)
        {
            this.Application.WindowSelectionChange -= eventHandler.OnWindowSelectionChange;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>WindowSize</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IWindowSizeEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.WindowSize += eventHandler.OnWindowSize;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>WindowSize</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IWindowSizeEventHandler eventHandler)
        {
            this.Application.WindowSize -= eventHandler.OnWindowSize;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>XMLSelectionChange</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IXMLSelectionChangeEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.XMLSelectionChange += eventHandler.OnXMLSelectionChange;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>XMLSelectionChange</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IXMLSelectionChangeEventHandler eventHandler)
        {
            this.Application.XMLSelectionChange -= eventHandler.OnXMLSelectionChange;
        }

        /// <summary>
        /// Subscribes the specified <i>Event Handler</i> to the <c>XMLValidationError</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to subscribe.</param>
        private void SubscribeEventHandler(IXMLValidationErrorEventHandler eventHandler)
        {
            this.UnsubscribeEventHandler(eventHandler);
            this.Application.XMLValidationError += eventHandler.OnXMLValidationError;
        }

        /// <summary>
        /// Unsubscribes the specified <i>Event Handler</i> from the <c>XMLValidationError</c> event.
        /// </summary>
        /// <param name="eventHandler">The <i>Event Handler</i> to unsubscribe.</param>
        private void UnsubscribeEventHandler(IXMLValidationErrorEventHandler eventHandler)
        {
            this.Application.XMLValidationError -= eventHandler.OnXMLValidationError;
        }
    }
}
