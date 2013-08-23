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
    /// The class <see cref="CustomDocumentPropertyWriter"/> represents a writer that can write the custom document
    /// properties in a <see cref="Word.Document"/>.
    /// </summary>
    public class CustomDocumentPropertyWriter
    {
        /// <summary>
        /// The custom document properties of a <see cref="Word.Document"/>.
        /// </summary>
        private readonly Core.DocumentProperties properties;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomDocumentPropertyWriter"/> class for the specified <see
        /// cref="Word.Document"/>.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> to write custom document properties to.</param>
        public CustomDocumentPropertyWriter(Word.Document document)
        {
            // TODO Some COM interfaces are "lately bound", therefore this won't work outside of a VSTO context.
            // Solutions that do use Reflection exist:
            // http://xtractpro.com/articles/Office-Properties.aspx?page=2
            // http://support.microsoft.com/kb/303296
            this.properties = (Core.DocumentProperties)document.CustomDocumentProperties;
        }

        /// <summary>
        /// Sets the value of the custom document property with the specified name to the specified value.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property to set.</param>
        /// <param name="propertyValue">The (new) value of the custom document property.</param>
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
        /// Deletes the custom document property with the specified name.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property to delete.</param>
        public void Delete(string propertyName)
        {
            this.properties[propertyName].Delete();
        }

        /// <summary>
        /// Adds a custom document property.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void AddProperty(
            string propertyName,
            object propertyValue,
            bool linkToContent = false,
            object linkSource = null)
        {
            // TODO Improve algorithm.
            if (propertyValue is string)
            {
                this.Set(propertyName, (string)propertyValue, linkToContent, linkSource);
            }
            else if (propertyValue is int)
            {
                this.Set(propertyName, (int)propertyValue, linkToContent, linkSource);
            }
            else if (propertyValue is float)
            {
                this.Set(propertyName, (float)propertyValue, linkToContent, linkSource);
            }
            else if (propertyValue is DateTime)
            {
                this.Set(propertyName, (DateTime)propertyValue, linkToContent, linkSource);
            }
            else if (propertyValue is bool)
            {
                this.Set(propertyName, (bool)propertyValue, linkToContent, linkSource);
            }
        }

        /// <summary>
        /// Adds a custom document property of type <c>string</c>.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void Set(string propertyName, string propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeString,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        /// <summary>
        /// Adds a custom document property of type <c>bool</c>.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void Set(string propertyName, bool? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeBoolean,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        /// <summary>
        /// Adds a custom document property of type <c>float</c>.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void Set(string propertyName, float? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeFloat,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        /// <summary>
        /// Adds a custom document property of type <c>int</c>.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void Set(string propertyName, int? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeNumber,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        /// <summary>
        /// Adds a custom document property of type <c>DateTime</c>.
        /// </summary>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void Set(string propertyName, DateTime? propertyValue = null, bool linkToContent = false, object linkSource = null)
        {
            this.Set(
                Core.MsoDocProperties.msoPropertyTypeDate,
                propertyName,
                propertyValue,
                linkToContent,
                linkSource);
        }

        /// <summary>
        /// Adds a custom document property of the specified data type.
        /// </summary>
        /// <param name="propertyType">The data type of the custom document property.</param>
        /// <param name="propertyName">The name of the custom document property.</param>
        /// <param name="propertyValue">The value of the custom document property.</param>
        /// <param name="linkToContent">
        /// Whether the custom document property is linked to the contents of the container document.
        /// </param>
        /// <param name="linkSource">
        /// The source of the linked property. Ignored if <c>linkToContent</c> is <c>false</c>.
        /// </param>
        private void Set(
            Core.MsoDocProperties propertyType,
            string propertyName,
            object propertyValue = null,
            bool linkToContent = false,
            object linkSource = null)
        {
            this.properties.Add(propertyName, linkToContent, propertyType, propertyValue, linkSource);
        }
    }
}
