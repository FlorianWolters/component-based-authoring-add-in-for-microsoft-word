//------------------------------------------------------------------------------
// <copyright file="CustomDocumentPropertyWriter.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.DocumentProperties
{
    using System;
    using Core = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="CustomDocumentPropertyWriter"/> represents a writer
    /// that can Set custom document properties to a Microsoft Word document.
    /// </summary>
    public class CustomDocumentPropertyWriter
    {
        /// <summary>
        /// The custom document properties of the Microsoft Word document.
        /// </summary>
        private readonly Core.DocumentProperties properties;

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="CustomDocumentPropertyWriter"/> class with the specified
        /// Microsoft Word document.
        /// </summary>
        /// <param name="document">The Microsoft Word document.</param>
        public CustomDocumentPropertyWriter(Word.Document document)
        {
            // TODO Some COM interfaces are "lately bound", therefore this won't
            // work outside of a VSTO context. Solutions that do use Reflection
            // exist:
            // http://xtractpro.com/articles/Office-Properties.aspx?page=2
            // http://support.microsoft.com/kb/303296
            this.properties = (Core.DocumentProperties)document.CustomDocumentProperties;
        }

        /// <summary>
        /// Sets the value of the custom property with the specified name to the
        /// specified value in the Microsoft Word document.
        /// </summary>
        /// <param name="propertyName">The name of the property to set.</param>
        /// <param name="propertyValue">The (new) value of the property.</param>
        public void Set(string propertyName, object propertyValue)
        {
            try
            {
                this.properties[propertyName].Value = propertyValue;
            }
            catch (ArgumentException)
            {
                this.AddProperty(propertyName, propertyValue);
            }
        }

        /// <summary>
        /// Deletes the custom property with the specified name from the
        /// Microsoft Word document.
        /// </summary>
        /// <param name="propertyName">The name of the property to delete.</param>
        public void Delete(string propertyName)
        {
            this.properties[propertyName].Delete();
        }

        private void AddProperty(
            string propertyName,
            object propertyValue,
            bool linkToContent = false,
            object linkSource = null)
        {
            // TODO Improve algorithm.
            if (propertyValue is string)
            {
                this.Set(
                    propertyName,
                    (string)propertyValue,
                    linkToContent,
                    linkSource);
            }
            else if (propertyValue is int)
            {
                this.Set(
                    propertyName,
                    (int)propertyValue,
                    linkToContent,
                    linkSource);
            }
            else if (propertyValue is float)
            {
                this.Set(
                    propertyName,
                    (float)propertyValue,
                    linkToContent,
                    linkSource);
            }
            else if (propertyValue is DateTime)
            {
                this.Set(
                    propertyName,
                    (DateTime)propertyValue,
                    linkToContent,
                    linkSource);
            }
            else if (propertyValue is bool)
            {
                this.Set(
                    propertyName,
                    (bool)propertyValue,
                    linkToContent,
                    linkSource);
            }
        }

        private void Set(string propertyName, string propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeString,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        private void Set(string propertyName, bool? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeBoolean,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        private void Set(string propertyName, float? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeFloat,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        private void Set(string propertyName, int? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeNumber,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        private void Set(string propertyName, DateTime? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeDate,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        private void Set(
            Core.MsoDocProperties propertyType,
            string propertyName,
            object propertyValue = null,
            bool linkToContent = false,
            object linkSource = null)
        {
            this.properties.Add(
                propertyName,
                linkToContent,
                propertyType,
                propertyValue,
                linkSource);
        }
    }
}
