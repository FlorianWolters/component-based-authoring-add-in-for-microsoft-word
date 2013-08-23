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

        /// <summary>
        /// Displays a message box in front of the specified object and with the message of the specified <see
        /// cref="Exception"/>.
        /// </summary>
        /// <param name="ex">The <see cref="Exception"/> whose message to display in the message box.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxWithExceptionMessage(Exception ex, IWin32Window owner = null)
        {
            return MessageBox.Show(owner, ex.Message, TitleForError, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Displays a message box in front of the specified object with a message that the file with the specified file
        /// path is read-only.
        /// </summary>
        /// <param name="filePath">The file path of the read-only file.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxFileIsReadOnly(string filePath, IWin32Window owner = null)
        {
             return MessageBox.Show(
                owner,
                "The file \"" + filePath + "\" is read-only and cannot be updated.",
                TitleForWarning,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Displays a message box in front of the specified object with a message that the files with the specified
        /// list of file paths are read-only.
        /// </summary>
        /// <param name="filePaths">The file paths of the read-only files.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
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

        /// <summary>
        /// Displays a message box in front of the specified object with a message that the file with the specified
        /// file path does not exist.
        /// </summary>
        /// <param name="filePath">The file path of the non-existing file.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxHelpFieldDoesNotExist(string filePath, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The help file \"" + filePath + "\" does not exist.",
                TitleForError,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePaths">The file paths of the files to update to.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePaths">The file paths of the files to update from.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxWhetherToOverwriteCustomDocumentProperty(
            string propertyName,
            IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "A custom document property with the name '" + propertyName + "' does already exist. Do you want to overwrite the value of the property?",
                TitleForQuestion,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxSetCustomDocumentPropertySuccess(string propertyName, string propertyValue, IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The custom document property with the name '" + propertyName + "' has been set to the value '" + propertyValue + "'.",
                TitleForInformation,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxNoCustomDocumentPropertyModfied(IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "The custom document properties of this document have not been modified.",
                TitleForQuestion,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
        public static DialogResult ShowMessageBoxInvalidFieldCodeFormat(IWin32Window owner = null)
        {
            return MessageBox.Show(
                owner,
                "Unable to parse the code of the field. Ensure that the code has the correct format.",
                TitleForWarning,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="owner">
        /// An implementation of <see cref="IWin32Window"/> that will own the modal dialog box.
        /// </param>
        /// <returns>One of the <see cref="DialogResult"/> values.</returns>
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
