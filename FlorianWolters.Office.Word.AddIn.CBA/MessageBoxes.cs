//------------------------------------------------------------------------------
// <copyright file="MessageBoxes.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA
{
    using System.Windows.Forms;

    public static class MessageBoxes
    {
        public static DialogResult ShowMessageBoxFileIsReadOnly(string filePath)
        {
            return MessageBox.Show(
                "The file \"" + filePath + "\" is read-only.",
                "Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public static DialogResult ShowMessageBoxHelpFieldDoesNotExist(string filePath)
        {
            return MessageBox.Show(
                "The help file \"" + filePath + "\" does not exist.",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        public static DialogResult ShowMessageBoxWhetherToUpdateContentInSource(int fieldCount)
        {
            return MessageBox.Show(
                "Do you really want to update the content in the " + fieldCount + " selected source file(s) with the content in this document?",
                "Question",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        public static DialogResult ShowMessageBoxWhetherToUpdateContentFromSource(int fieldCount)
        {
            return MessageBox.Show(
                "Do you really want to update the content in this document with the content of the " + fieldCount + " selected source file(s)?",
                "Question",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        public static DialogResult ShowMessageBoxWhetherToOverwriteCustomDocumentProperty(string propertyName)
        {
            return MessageBox.Show(
                "A custom document property with the name '" + propertyName + "' does already exist. Do you want to overwrite the value of the property?",
                "Question",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
        }

        public static DialogResult ShowMessageBoxSetCustomDocumentPropertySuccess(string propertyName, string propertyValue)
        {
            return MessageBox.Show(
                "The custom document property with the name '" + propertyName + "' has been set to the value '" + propertyValue + "'.",
                "Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public static DialogResult ShowMessageBoxNoCustomDocumentPropertyModfied()
        {
            return MessageBox.Show(
                "The custom document properties of this document have not been modified.",
                "Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
