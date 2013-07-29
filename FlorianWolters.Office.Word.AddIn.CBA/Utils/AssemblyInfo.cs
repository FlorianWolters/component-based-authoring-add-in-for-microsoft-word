//------------------------------------------------------------------------------
// <copyright file="AssemblyInfo.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Utils
{
    using System;
    using System.IO;
    using System.Reflection;

    public class AssemblyInfo
    {
        private readonly Assembly assembly;

        public AssemblyInfo(Assembly assembly)
        {
            if (null == assembly)
            {
                throw new ArgumentNullException("assembly");
            }

            this.assembly = assembly;
        }

        public string Company
        {
            get
            {
                return ((AssemblyCompanyAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyCompanyAttribute), false))
                   .Company;
            }
        }

        public string Copyright
        {
            get
            {
                return ((AssemblyCopyrightAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyCopyrightAttribute), false))
                   .Copyright;
            }
        }

        public string Description
        {
            get
            {
                return ((AssemblyDescriptionAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyDescriptionAttribute), false))
                    .Description;
            }
        }

        public string Product
        {
            get
            {
                return ((AssemblyProductAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyProductAttribute), false))
                    .Product;
            }
        }

        public string Title
        {
            get
            {
                return ((AssemblyTitleAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyTitleAttribute), false))
                    .Title;
            } 
        }

        public Version Version
        {
            get
            {
                return this.assembly.GetName().Version;
            }
        }

        public string CodeBasePath
        {
            get
            {
                return Path.GetDirectoryName(
                    new Uri(this.assembly.CodeBase).LocalPath);
            }
        }
    }
}
