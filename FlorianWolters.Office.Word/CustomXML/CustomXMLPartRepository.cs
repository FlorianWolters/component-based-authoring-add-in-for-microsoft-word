//------------------------------------------------------------------------------
// <copyright file="CustomXMLPartRepository.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.CustomXML
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml;
    using System.Xml.Linq;
    using Office = Microsoft.Office.Core;

    /// <summary>
    /// The class <see cref="CustomXMLPartRepository"/> allows to manage <see cref="Office.CustomXMLParts"/> via default
    /// namespaces.
    /// </summary>
    public class CustomXMLPartRepository
    {
        /// <summary>
        /// The <see cref="Office.CustomXMLParts"/> of an Microsoft Word <see cref="Word.Document"/>.
        /// </summary>
        private readonly Office.CustomXMLParts customXMLParts;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomXMLPartRepository"/> class for the specified <see
        /// cref="Office.CustomXMLParts"/>.
        /// </summary>
        /// <param name="customXMLParts">The <see cref="Office.CustomXMLParts"/>.</param>
        public CustomXMLPartRepository(Office.CustomXMLParts customXMLParts)
        {
            this.customXMLParts = customXMLParts;
        }

        /// <summary>
        /// Returns all <see cref="Office.CustomXMLPart"/>s.
        /// </summary>
        /// <returns>All <see cref="Office.CustomXMLPart"/>s.</returns>
        public IEnumerable<Office.CustomXMLPart> GetAll()
        {
            return from Office.CustomXMLPart x in this.customXMLParts
                   select x;
        }

        /// <summary>
        /// Returns the <see cref="Office.CustomXMLPart"/> with the specified  default namespace.
        /// </summary>
        /// <param name="defaultNamespace">
        /// The default namespace of the <see cref="Office.CustomXMLPart"/> to return.
        /// </param>
        /// <returns>The <see cref="Office.CustomXMLPart"/> with the specified default namespace.</returns>
        public Office.CustomXMLPart Get(string defaultNamespace)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            this.ThrowExceptionIfNamespaceNotExists(defaultNamespace);

            return this.FindByDefaultNamespace(defaultNamespace);
        }

        /// <summary>
        /// Determines whether one <see cref="Office.CustomXMLPart"/> with the specified default namespace exists.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <returns>
        /// <c>true</c> if one <see cref="Office.CustomXMLPart"/> with the specified default namespace exists;
        /// <c>false</c> otherwise.
        /// </returns>
        public bool Exists(string defaultNamespace)
        {
            return null != this.FindByDefaultNamespace(defaultNamespace);
        }

        /// <summary>
        /// Adds a new <see cref="Office.CustomXMLPart"/> with the specified default namespace and the specified XML
        /// data.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <param name="xml">The XML data.</param>
        /// <returns>The newly created <see cref="Office.CustomXMLPart"/>.</returns>
        public Office.CustomXMLPart Add(string defaultNamespace, string xml)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            this.ThrowExceptionIfNamespaceExists(defaultNamespace);

            return this.customXMLParts.Add(xml);
        }

        /// <summary>
        /// Deletes the <see cref="Office.CustomXMLPart"/> with the specified default namespace.
        /// </summary>
        /// <param name="defaultNamespace">
        /// The default namespace of the <see cref="Office.CustomXMLPart"/> to delete.
        /// </param>
        public void Delete(string defaultNamespace)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            this.ThrowExceptionIfNamespaceNotExists(defaultNamespace);

            this.FindByDefaultNamespace(defaultNamespace).Delete();
        }

        /// <summary>
        /// Returns the <see cref="Office.CustomXMLPart"/> with the specified default namespace.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <returns>
        /// The <see cref="Office.CustomXMLPart"/> for the specified default namespace on success; <c>null</c> if no
        /// <see cref="Office.CustomXMLPart"/> with the specified default namespace exists.
        /// </returns>
        public Office.CustomXMLPart FindByDefaultNamespace(string defaultNamespace)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            Office.CustomXMLParts candidates = this.customXMLParts.SelectByNamespace(defaultNamespace);

            if (candidates.Count > 1)
            {
                this.ThrowExceptionIfNamespaceExists(defaultNamespace);
            }

            return (1 == candidates.Count)
                ? candidates[1]
                : null;
        }

        /// <summary>
        /// Returns the <see cref="Office.CustomXMLPart"/> with the specified Globally Unique Identifier (GUID). 
        /// </summary>
        /// <param name="guid">The Globally Unique Identifier (GUID).</param>
        /// <returns>
        /// The <see cref="Office.CustomXMLPart"/> with the specified GUID on success; <c>null</c> if no <see
        /// cref="Office.CustomXMLPart"/> with the specified GUID exists.
        /// </returns>
        public Office.CustomXMLPart FindByID(string guid)
        {
            return this.customXMLParts.SelectByID(guid);
        }

        /// <summary>
        /// Returns all built-in <see cref="Office.CustomXMLPart"/>s.
        /// </summary>
        /// <returns>All built-in <see cref="Office.CustomXMLPart"/>s.</returns>
        public IEnumerable<Office.CustomXMLPart> FindBuiltIn()
        {
            return from Office.CustomXMLPart x in this.customXMLParts
                   where x.BuiltIn
                   select x;
        }

        /// <summary>
        /// Returns all not built-in <see cref="Office.CustomXMLPart"/>s.
        /// </summary>
        /// <returns>All not built-in <see cref="Office.CustomXMLPart"/>s.</returns>
        public IEnumerable<Office.CustomXMLPart> FindNotBuiltIn()
        {
            return from Office.CustomXMLPart x in this.customXMLParts
                   where !x.BuiltIn
                   select x;
        }

        /// <summary>
        /// Deletes each <see cref="Office.CustomXMLPart"/> which is not built-in.
        /// </summary>
        public void DeleteAllNotBuiltIn()
        {
            this.FindNotBuiltIn().ToList().ForEach(x => x.Delete());
        }

        /// <summary>
        /// Synchronizes the XML files in the specified directory path with the <see cref="Office.CustomXMLParts"/>.
        /// </summary>
        /// <param name="directoryPath">The path of the directory to synchronize.</param>
        public void SynchronizeWithDirectory(string directoryPath)
        {
            string[] filePaths = this.XMLFilePathsFromDirectoryPath(directoryPath);
            string defaultNamespace = null;

            // TODO This is too strict. We do have to think about the interopability! A better approach would be to
            // allow the user to delete a custom XML part.
            // Retrieve and delete all custom XML parts that have a default namespace which does not exist in any of the
            // XML files.
            ////(from c in this.FindNotBuiltIn()
            //// where !fileDefaultNamespaces.Contains(c.NamespaceURI)
            //// select c).ToList().ForEach(c => c.Delete());

            foreach (string filePath in filePaths)
            {
                defaultNamespace = this.DefaultNamespaceFromXMLFile(filePath);

                if (this.Exists(defaultNamespace))
                {
                    // Deleting and adding a custom XML part with the same namespaces seems to work (the bindings are
                    // not lost). Other solutions are much more complicated.
                    this.Delete(defaultNamespace);
                }

                this.AddFromFile(filePath);
            }
        }

        /// <summary>
        /// Adds a new <see cref="Word.CustomXMLPart"/> for each XML file in the specified directory.
        /// </summary>
        /// <param name="directoryPath">The path of the directory.</param>
        public void AddFromDirectory(string directoryPath)
        {
            this.XMLFilePathsFromDirectoryPath(directoryPath)
                .ToList()
                .ForEach(filePath => this.AddFromFile(filePath));
        }

        /// <summary>
        /// Adds a new <see cref="Word.CustomXMLPart"/> for the specified XML file.
        /// </summary>
        /// <param name="filePath">The path of the XML file.</param>
        /// <returns>The newly created <see cref="Word.CustomXMLPart"/>.</returns>
        public Office.CustomXMLPart AddFromFile(string filePath)
        {
            Office.CustomXMLPart result = null;
            string defaultNamespace = this.DefaultNamespaceFromXMLFile(filePath);

            try
            {
                // The Load method of class Office.CustomXMLPart is broken, since it can't handle return characters
                // properly, e.g. the Office.CustomXMLPart object created from a XML file with ONE child node (of the
                // root node) contains THREE child nodes. Therefore we use class XmlDocument to read the XML.
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(filePath);
                result = this.Add(defaultNamespace, xmlDocument.OuterXml);
            }
            catch (ArgumentException)
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "The XML file \"" + filePath + "\" does not have a default namespace declaration.");
            }
            catch (CustomXMLPartDefaultNamespaceException)
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "The XML file \"" + filePath + "\" does have the same default namespace declaration (\""
                    + defaultNamespace + "\") as an already existing custom XML part.");
            }

            return result;
        }

        /// <summary>
        /// Returns the default namespace for the specified XML file.
        /// </summary>
        /// <param name="filePath">The path of the XML file.</param>
        /// <returns>The default namespace of the XML file.</returns>
        private string DefaultNamespaceFromXMLFile(string filePath)
        {
            XElement element = XElement.Load(filePath);

            return element.GetDefaultNamespace().NamespaceName;
        }

        /// <summary>
        /// Returns the file paths of the XML files in the specified directory.
        /// </summary>
        /// <param name="directoryPath">The path of the directory.</param>
        /// <returns>The the file paths of the XML files.</returns>
        private string[] XMLFilePathsFromDirectoryPath(string directoryPath)
        {
            return Directory.GetFiles(directoryPath, "*.xml", SearchOption.AllDirectories);
        }

        /// <summary>
        /// Validates the specified default namespace.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace to validate.</param>
        private void ValidateDefaultNamespace(string defaultNamespace)
        {
            this.ThrowExceptionIfNamespaceIsNull(defaultNamespace);
            this.ThrowExceptionIfNamespaceIsEmpty(defaultNamespace);
        }

        /// <summary>
        /// Throws an <see cref="ArgumentNullException"/> if the specified default namespace is <c>null</c>.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <exception cref="ArgumentNullException">If <c>defaultNamespace</c> is <c>null</c>.</exception>
        private void ThrowExceptionIfNamespaceIsNull(string defaultNamespace)
        {
            if (null == defaultNamespace)
            {
                throw new ArgumentNullException("defaultNamespace");
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if the specified default namespace is empty.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <exception cref="ArgumentException">If <c>defaultNamespace</c> is <c>""</c>.</exception>
        private void ThrowExceptionIfNamespaceIsEmpty(string defaultNamespace)
        {
            if (string.Empty == defaultNamespace)
            {
                throw new ArgumentException(
                    "defaultNamespace cannot be an empty string",
                    "defaultNamespace");
            }
        }

        /// <summary>
        /// Throws a <see cref="CustomXMLPartDefaultNamespaceException"/> if the specified default namespace does
        /// already exist.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <exception cref="CustomXMLPartDefaultNamespaceException">
        /// If <c>defaultNamespace</c> does already exist.
        /// </exception>
        private void ThrowExceptionIfNamespaceExists(string defaultNamespace)
        {
            if (this.Exists(defaultNamespace))
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "A custom XML part with the default namespace \""
                    + defaultNamespace + "\" does already exist.");
            }
        }

        /// <summary>
        /// Throws a <see cref="CustomXMLPartDefaultNamespaceException"/> if the specified default namespace does not
        /// exist.
        /// </summary>
        /// <param name="defaultNamespace">The default namespace.</param>
        /// <exception cref="CustomXMLPartDefaultNamespaceException">
        /// If <c>defaultNamespace</c> does not exist.
        /// </exception>
        private void ThrowExceptionIfNamespaceNotExists(string defaultNamespace)
        {
            if (!this.Exists(defaultNamespace))
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "A custom XML part with the default namespace \""
                    + defaultNamespace + "\" does not exist.");
            }
        }
    }
}
