//------------------------------------------------------------------------------
// <copyright file="MarkdownForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Text;
    using System.Windows.Forms;
    using MarkdownSharp;

    /// <summary>
    /// The class <see cref="MarkdownForm"/> implements a simple modal window which allows to display a file that
    /// contains Markdown-formatted content.
    /// </summary>
    internal partial class MarkdownForm : Form
    {
        /// <summary>
        /// Transform the Markdown-formatted text of the file to HTML.
        /// </summary>
        private readonly Markdown markdown;

        /// <summary>
        /// Determines whether the file has been opened.
        /// <remarks>
        /// This field is used to prevent the opening of links in the <see cref="WebBrowser"/> of this <see
        /// cref="MarkdownForm"/>.
        /// </remarks>
        /// </summary>
        private bool hasNavigated = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="MarkdownForm"/> class with the specified file to display.
        /// </summary>
        /// <param name="filePath">The file path of the document to display.</param>
        public MarkdownForm(string filePath)
            : this(filePath, new Markdown())
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MarkdownForm"/> class with the specified file to display and
        /// the specified <see cref="Markdown"/> instance.
        /// </summary>
        /// <param name="filePath">The file path of the document to display.</param>
        /// <param name="markdown">Transform the Markdown-formatted text of the file to HTML.</param>
        public MarkdownForm(string filePath, Markdown markdown)
            : this()
        {
            this.markdown = markdown;
            this.ChangeDocument(filePath);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MarkdownForm"/> class.
        /// </summary>
        public MarkdownForm()
        {
            this.InitializeComponent();
            this.markdown = new Markdown();
            this.webBrowser.Navigating += this.OnNavigating;
        }

        /// <summary>
        /// Changes the document to display in this <see cref="MarkdownForm"/> to the specified file path.
        /// </summary>
        /// <param name="filePath">The file path of the document to display.</param>
        public void ChangeDocument(string filePath)
        {
            string markdownText = File.ReadAllText(filePath);
            string htmlText = this.markdown.Transform(markdownText);

            StringBuilder stringBuilder = new StringBuilder("<style type=\"text/css\">");
            stringBuilder.Append(Environment.NewLine);
            stringBuilder.Append("* { font-family: Arial, \"Helvetica Neue\", Helvetica, sans-serif; }");
            stringBuilder.Append(Environment.NewLine);
            stringBuilder.Append("</style>");
            stringBuilder.Append(Environment.NewLine);
            stringBuilder.Append(htmlText);
            this.webBrowser.DocumentText = stringBuilder.ToString();
        }

        /// <summary>
        /// Changes the title of this <see cref="MarkdownForm"/>.
        /// </summary>
        /// <param name="text">The new title.</param>
        public void ChangeTitle(string text)
        {
            this.Text = text;
        }

        /// <summary>
        /// Raises the <c>Navigating</c> event.
        /// <para>
        /// Prevents the opening of links in the <see cref="WebBrowser"/> of this <see cref="MarkdownForm"/>. Instead
        /// links are opened with the default program.
        /// </para>
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A <see cref="WebBrowserNavigatingEventArgs"/> that contains the event data.</param>
        /// <remarks>
        /// The source code has been taken from <a
        /// href="http://stackoverflow.com/questions/10927696/webbrowser-control-open-default-browser-on-click-on-link">this</a>
        /// Stack Overflow question.
        /// </remarks>
        private void OnNavigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (!this.hasNavigated)
            {
                this.hasNavigated = true;
                return;
            }

            e.Cancel = true;

            Process.Start(e.Url.ToString());
        }
    }
}
