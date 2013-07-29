//------------------------------------------------------------------------------
// <copyright file="CustomDocumentPropertyReader.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.DocumentProperties
{
    using System.Collections.Generic;
    using System.Linq;
    using Core = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="CustomDocumentPropertyReader"/> represents a reader
    /// that can read custom document properties from a Microsoft Word document.
    /// </summary>
    public class CustomDocumentPropertyReader
    {
        /// <summary>
        /// The character used as a prefix to mark <i>internal</i> custom
        /// document properties.
        /// </summary>
        public const char InternalPrefixCharacter = '_';

        /// <summary>
        /// The custom document properties of the Microsoft Word document.
        /// </summary>
        private Core.DocumentProperties properties;

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="CustomDocumentPropertyReader"/> class with the specified
        /// Microsoft Word document.
        /// </summary>
        /// <param name="document">The Microsoft Word document to load.</param>
        public CustomDocumentPropertyReader(Word.Document document) : this()
        {
            this.Load(document);
        }

        /// <summary>
        /// Initializes a new instance of the <see
        /// cref="CustomDocumentPropertyReader"/> class.
        /// </summary>
        public CustomDocumentPropertyReader()
        {
        }

        /// <summary>
        /// Loads the custom document properties from the specified Microsoft
        /// Word document.
        /// </summary>
        /// <param name="document">The Microsoft Word document to load.</param>
        public void Load(Word.Document document)
        {
            // TODO Some COM interfaces are "lately bound", therefore this won't
            // work outside of a VSTO context. Solutions that do use Reflection
            // exist:
            // http://xtractpro.com/articles/Office-Properties.aspx?page=2
            // http://support.microsoft.com/kb/303296
            this.properties = (Core.DocumentProperties)document.CustomDocumentProperties;
        }

        /// <summary>
        /// Returns all custom document properties from the Microsoft Word
        /// document, loaded with this <see
        /// cref="CustomDocumentPropertyReader"/>.
        /// </summary>
        /// <returns>All custom document properties.</returns>
        public IEnumerable<Core.DocumentProperty> GetAll()
        {
            return this.properties.Cast<Core.DocumentProperty>();
        }

        /// <summary>
        /// Returns all <i>internal</i> custom document properties from the
        /// Microsoft Word document, loaded with this <see
        /// cref="CustomDocumentPropertyReader"/>.
        /// <para>
        /// <i>Internal</i> custom document properties start with the character
        /// <c>_</c>.
        /// </para>
        /// </summary>
        /// <returns>All <i>internal</i> custom document properties.</returns>
        public IEnumerable<Core.DocumentProperty> FindAllExceptInternal()
        {
            return from property in this.GetAll()
                   where InternalPrefixCharacter != property.Name[0]
                   select property;
        }

        /// <summary>
        /// Determines whether a custom document property with the specified
        /// name exists in the Microsoft Word document.
        /// </summary>
        /// <param name="propertyName">The name of the property to check.</param>
        /// <returns><c>true</c> if the property with the specified name exists, <c>false</c> otherwise.</returns>
        public bool Exists(string propertyName)
        {
            return 1 == this.properties.Cast<Core.DocumentProperty>()
                .Where(c => c.Name == propertyName)
                .Count();
        }

        /// <summary>
        /// Returns the value of the custom property with the specified name
        /// from the Microsoft Word document and casts it to the specified type.
        /// </summary>
        /// <typeparam name="T">
        /// The type to cast the value of the custom document property to.
        /// </typeparam>
        /// <param name="propertyName">The name of the property to read.</param>
        /// <returns>The value of the property with the specified name.</returns>
        public T Get<T>(string propertyName) where T : struct
        {
            return (T)this.Get(propertyName);
        }

        /// <summary>
        /// Returns the value of the custom property with the specified name
        /// from the Microsoft Word document.
        /// </summary>
        /// <param name="propertyName">The name of the property to read.</param>
        /// <returns>The value of the property with the specified name.</returns>
        public object Get(string propertyName)
        {
            if (!this.Exists(propertyName))
            {
                throw new UnknownCustomDocumentPropertyException(
                    "The custom Document property \"" + propertyName + "\" does not exist in the Microsoft Word Document.");
            }

            return this.properties[propertyName].Value;
        }
    }
}
