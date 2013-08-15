//------------------------------------------------------------------------------
// <copyright file="MessageBoxes.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Windows.Forms;

    /// <summary>
    /// The static class <see cref="MessageBoxes"/> implements default message boxes for the application.
    /// </summary>
    public static class MessageBoxes
    {
        /// <summary>
        /// The text to display in the title bar of the message box for an error.
        /// </summary>
        private const string TitleForError = "Error";

        /// <summary>
        /// The text to display in the title bar of the message box for an information.
        /// </summary>
        private const string TitleForInformation = "Information";

        /// <summary>
        /// The text to display in the title bar of the message box for a question.
        /// </summary>
        private const string TitleForQuestion = "Question";

        /// <summary>
        /// The text to display in the title bar of the message box for a warning.
        /// </summary>
        private const string TitleForWarning = "Warning";

        public static DialogResult ShowMessageBoxWithExceptionMessage(Exception ex, IWin32Window owner = null)
        {
            return MessageBox.Show(owner, ex.Message, TitleForError, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static DialogResult ShowMessageBoxFileIsReadOnly(string filePath, IWin32Window owner = null)
        {
            return MessageBox.Show(owner,
                "The file \"" + filePath + "\" is read-only and cannot be updated.",
                TitleForWarning,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public static DialogResult ShowMessageBoxFileIsReadOnly(IList<string> filePaths, IWin32Window owner = null)
        {
            StringBuilder message = new StringBuilder("The following ");
            message.Append(filePaths.Count);
            message.Append(" referenced source file");
            
            if (filePaths.Count > 1)
            {
                message.Append("s are");
            }
            else
            {
                message.Append(" is");
            }

            message.Append(" read-only:");
            message.Append(Environment.NewLine);
            message.Append(string.Join(Environment.NewLine, filePaths));

            return MessageBox.Show(
                owner,
                message.ToString(),
                TitleForInformation,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public static DialogResult ShowMessageBoxHelpFieldDoesNotExist(string filePath, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The help file \"" + filePath + "\" does not exist.",
                TitleForError,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        public static DialogResult ShowMessageBoxWhetherToUpdateContentInSource(IList<string> filePaths, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "Do you really want to update the content in the following " + filePaths.Count + " referenced source file(s) with the content from this document?" + Environment.NewLine + string.Join(Environment.NewLine, filePaths),
                TitleForQuestion,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        public static DialogResult ShowMessageBoxWhetherToUpdateContentFromSource(IList<string> filePaths, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "Do you really want to update the content in this document with the content from the following " + filePaths.Count + " referenced source file(s)?" + Environment.NewLine + string.Join(Environment.NewLine, filePaths),
                TitleForQuestion,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        public static DialogResult ShowMessageBoxWhetherToOverwriteCustomDocumentProperty(string propertyName, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "A custom document property with the name '" + propertyName + "' does already exist. Do you want to overwrite the value of the property?",
                TitleForQuestion,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        public static DialogResult ShowMessageBoxSetCustomDocumentPropertySuccess(string propertyName, string propertyValue, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The custom document property with the name '" + propertyName + "' has been set to the value '" + propertyValue + "'.",
                TitleForInformation,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public static DialogResult ShowMessageBoxNoCustomDocumentPropertyModfied(IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The custom document properties of this document have not been modified.",
                TitleForQuestion,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public static DialogResult ShowMessageBoxInvalidFieldCodeFormat(IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "Unable to parse the code of the field. Ensure that the code has the correct format.",
                TitleForWarning,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public static DialogResult ShowMessageBoxFieldResultIsEqualToSourceFile(IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The content in this document and the content of the referenced source file are equal.",
                TitleForInformation,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
