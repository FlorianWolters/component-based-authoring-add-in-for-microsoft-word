//------------------------------------------------------------------------------
// <copyright file="ThisAddIn.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA
{
    using System;
    using System.Globalization;
    using System.Threading;
    using Microsoft.Office.Tools;
    using Office = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="ThisAddIn"/> is the entry point of the <i>Microsoft
    /// Word</i> Application-Level Add-in.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Returns an object that implements the <see
        /// ref="Office.IRibbonExtensibility"/> interface.
        /// <para>
        /// Sets the <see cref="CultureInfo"/> for the <i>Ribbons</i> of <see
        /// cref="ThisAddIn"/> to the language of the <i>Microsoft Word</i>
        /// applicationEvent.
        /// </para>
        /// </summary>
        /// <returns>The extension for the <i>Ribbons</i>.</returns>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            // We can't access the property "this.Application" here, since this
            // method is invoked before the event handler, triggered by the
            // "this.Startup" event. Therefore "this.Application" is "null" here
            // and we do have to retrieve the
            // "Microsoft.Office.Interop.Word.Application" otherwise.
            this.ChangeCultureOfCurrentThreadToCultureOfWordApplication(
                this.RetrieveWordApplication(this));

            return base.CreateRibbonExtensibilityObject();
        }

        /// <summary>
        /// Returns the <see cref="Word.Application"/> of <see
        /// cref="ThisAddIn"/>.
        /// </summary>
        /// <param name="addIn">An Application-Level Add-in.</param>
        /// <returns>The Word application.</returns>
        private Word.Application RetrieveWordApplication(AddInBase addIn)
        {
            return this.GetHostItem<Word.Application>(
                typeof(Word.Application),
                "Application");
        }

        /// <summary>
        /// Executes code when <see cref="ThisAddIn"/> is loaded, after all the
        /// initialization code in the assembly has run.
        /// </summary>
        /// <remarks>
        /// The <see cref="ThisAddIn_Startup"/> is a default event handler.
        /// </remarks>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The arguments of the event.</param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Attention: "VSTO" waits for "Microsoft Word" to be ready before
            // firing the "Startup" event. Therefore the "DocumentOpen" and
            // "WindowActivate" events may have already fired.

            // Sets the culture of this Add-In (e.g. dialogs) to the culture of
            // the "Microsoft Word" application.
            this.ChangeCultureOfCurrentThreadToCultureOfWordApplication(
                this.Application);
        }

        /// <summary>
        /// Executes code when <see cref="ThisAddIn"/> is about to be unloaded. 
        /// </summary>
        /// <remarks>
        /// The <see cref="ThisAddIn_Startup"/> is a default event handler.
        /// </remarks>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">The arguments of the event.</param>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // NOOP
        }

        /// <summary>
        /// Sets the <see cref="CultureInfo"/> of <see cref="ThisAddIn"/> to the
        /// language of the specified <see cref="Word.Application"/>.
        /// </summary>
        /// <param name="application">A <see cref="Word.Application"/>.</param>
        private void ChangeCultureOfCurrentThreadToCultureOfWordApplication(
            Word.Application application)
        {
            int localeId = this.GetLocaleIdOfWordApplication(application);

            // TODO Implement I18N:
            // 1. Set the Localizable property of the Ribbon to true.
            // 2. Create dedicated resource file for each culture.
            // Uncomment to test with the "en" culture.
            // localeId = 1033;
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(localeId);
        }

        /// <summary>
        /// Returns the Locale ID (LCID) of the specified <see
        /// cref="Word.Application"/>.
        /// </summary>
        /// <remarks>
        /// <b>Locale ID (LCID)</b>: A 32-bit value defined by Microsoft Windows
        /// that consists of a language ID, sort ID, and reserved bits that
        /// identify a particular language. For example, the LCID for English is
        /// 1033, and the LCID for German is 1031. 
        /// </remarks>
        /// <param name="application">A <see cref="Word.Application"/>.</param>
        /// <returns>The LCID of the A <see cref="Word.Application"/>.</returns>
        private int GetLocaleIdOfWordApplication(Word.Application application)
        {
            return application.LanguageSettings.get_LanguageID(
                Office.MsoAppLanguageID.msoLanguageIDUI);
        }

        /// <summary>
        /// Required method for Designer support - do not modify the contents of
        /// this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += this.ThisAddIn_Startup;
            this.Shutdown += this.ThisAddIn_Shutdown;
        }
    }
}
