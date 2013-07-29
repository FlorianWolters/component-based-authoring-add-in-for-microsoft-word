//------------------------------------------------------------------------------
// <copyright file="CustomXMLPartRepository.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.CustomXML
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;
    using Office = Microsoft.Office.Core;

    // TODO Refactor class, since it does seem to have more than one responsibility.

    /// <summary>
    /// The class <see cref="CustomXMLPartRepository"/> allows to manage custom
    /// XML parts via default namespaces.
    /// </summary>
    public class CustomXMLPartRepository
    {
        private readonly Office.CustomXMLParts customXMLParts;

        public CustomXMLPartRepository(Office.CustomXMLParts customXMLParts)
        {
            this.customXMLParts = customXMLParts;
        }

        public IEnumerable<Office.CustomXMLPart> GetAll()
        {
            // TODO Possible to convert Office.CustomXMLParts to IEnumerable?
            IList<Office.CustomXMLPart> result = new List<Office.CustomXMLPart>();

            foreach (Office.CustomXMLPart part in this.customXMLParts)
            {
                result.Add(part);
            }

            return result;
        }

        public Office.CustomXMLPart Get(string defaultNamespace)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            this.ThrowExceptionIfNamespaceExists(defaultNamespace);

            return this.FindByDefaultNamespace(defaultNamespace)[0];
        }

        public bool Exists(string defaultNamespace)
        {
            return this.FindByDefaultNamespace(defaultNamespace).Count > 0;
        }

        public Office.CustomXMLPart Add(string defaultNamespace, string xml)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            this.ThrowExceptionIfNamespaceExists(defaultNamespace);

            return this.customXMLParts.Add(xml);
        }

        public void Delete(string defaultNamespace)
        {
            this.ValidateDefaultNamespace(defaultNamespace);
            this.ThrowExceptionIfNamespaceNotExists(defaultNamespace);

            this.FindByDefaultNamespace(defaultNamespace)[0].Delete();
        }

        public Office.CustomXMLPart FindByID(string id)
        {
            return this.customXMLParts.SelectByID(id);
        }

        public IEnumerable<Office.CustomXMLPart> FindBuiltIn()
        {
            IList<Office.CustomXMLPart> result = new List<Office.CustomXMLPart>();

            foreach (Office.CustomXMLPart part in this.customXMLParts)
            {
                if (part.BuiltIn)
                {
                    result.Add(part);
                }
            }

            return result;
        }

        public IEnumerable<Office.CustomXMLPart> FindNotBuiltIn()
        {
            IList<Office.CustomXMLPart> result = new List<Office.CustomXMLPart>();

            foreach (Office.CustomXMLPart part in this.customXMLParts)
            {
                if (!part.BuiltIn)
                {
                    result.Add(part);
                }
            }

            return result;
        }

        public void DeleteAllNotBuiltIn()
        {
            this.FindNotBuiltIn().ToList().ForEach(x => x.Delete());
        }

        public bool AddFromDirectory(string directoryPath)
        {
            bool result = false;

            if (Directory.Exists(directoryPath))
            {
                string[] filePaths = Directory.GetFiles(
                    directoryPath,
                    "*.xml",
                    SearchOption.AllDirectories);

                if (filePaths.Count() > 0)
                {
                    filePaths.ToList().ForEach(
                        f => this.AddFromFile(f));
                    result = true;
                }
            }

            return result;
        }

        public Office.CustomXMLPart AddFromFile(string filePath)
        {
            Office.CustomXMLPart result = null;
            string defaultNamespace = this.ReadDefaultNamespaceFromFile(filePath);

            try
            {
                result = this.Add(defaultNamespace, null);
                result.Load(filePath);
            }
            catch (ArgumentException)
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "The XML file \"" + filePath + "\" does not have a default namespace declaration.");
            }
            catch (CustomXMLPartDefaultNamespaceException)
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "The XML file \"" + filePath + "\" does have the same default namespace declaration as an already existing custom XML part.");
            }

            return result;
        }

        private string ReadDefaultNamespaceFromFile(string filePath)
        {
            XElement element = XElement.Load(filePath);

            return element.GetDefaultNamespace().NamespaceName;
        }

        private Office.CustomXMLParts FindByDefaultNamespace(string defaultNamespace)
        {
            return this.customXMLParts.SelectByNamespace(defaultNamespace);
        }

        private void ValidateDefaultNamespace(string defaultNamespace)
        {
            this.ThrowExceptionIfNamespaceIsNull(defaultNamespace);
            this.ThrowExceptionIfNamespaceIsEmpty(defaultNamespace);
        }

        private void ThrowExceptionIfNamespaceIsNull(string defaultNamespace)
        {
            if (null == defaultNamespace)
            {
                throw new ArgumentNullException("defaultNamespace cannot be null");
            }
        }

        private void ThrowExceptionIfNamespaceIsEmpty(string defaultNamespace)
        {
            if (string.Empty == defaultNamespace)
            {
                throw new ArgumentException("defaultNamespace cannot be an empty string");
            }
        }

        private void ThrowExceptionIfNamespaceExists(string defaultNamespace)
        {
            if (this.Exists(defaultNamespace))
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "A custom XML part with the specified default namespace does already exist.");
            }
        }

        private void ThrowExceptionIfNamespaceNotExists(string defaultNamespace)
        {
            if (!this.Exists(defaultNamespace))
            {
                throw new CustomXMLPartDefaultNamespaceException(
                    "A custom XML part with the specified default namespace does not exist.");
            }
        }
    }
}
