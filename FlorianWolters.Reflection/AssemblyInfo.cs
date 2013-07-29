//------------------------------------------------------------------------------
// <copyright file="AssemblyInfo.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Reflection
{
    using System;
    using System.IO;
    using System.Reflection;

    /// <summary>
    /// The class <see cref="AssemblyInfo"/> represents a reader that can read
    /// information from an <see cref="Assembly"/>.
    /// </summary>
    public class AssemblyInfo
    {
        /// <summary>
        /// The <see cref="Assembly"/> to read from.
        /// </summary>
        private readonly Assembly assembly;

        /// <summary>
        /// Initializes a new instance of the <see cref="AssemblyInfo"/> class
        /// with the specified <see cref="Assembly"/> to read from.
        /// </summary>
        /// <param name="assembly">The <see cref="Assembly"/> to read from.</param>
        /// <exception cref="ArgumentNullException">If the <c>assembly</c> argument is <c>null</c>.</exception>
        public AssemblyInfo(Assembly assembly)
        {
            if (null == assembly)
            {
                throw new ArgumentNullException("assembly");
            }

            this.assembly = assembly;
        }

        /// <summary>
        /// Gets the company from the <see cref="Assembly"/>.
        /// </summary>
        public string Company
        {
            get
            {
                return ((AssemblyCompanyAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyCompanyAttribute), false))
                   .Company;
            }
        }

        /// <summary>
        /// Gets the copyright from the <see cref="Assembly"/>.
        /// </summary>
        public string Copyright
        {
            get
            {
                return ((AssemblyCopyrightAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyCopyrightAttribute), false))
                   .Copyright;
            }
        }

        /// <summary>
        /// Gets the description from the <see cref="Assembly"/>.
        /// </summary>
        public string Description
        {
            get
            {
                return ((AssemblyDescriptionAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyDescriptionAttribute), false))
                    .Description;
            }
        }

        /// <summary>
        /// Gets the product from the <see cref="Assembly"/>.
        /// </summary>
        public string Product
        {
            get
            {
                return ((AssemblyProductAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyProductAttribute), false))
                    .Product;
            }
        }

        /// <summary>
        /// Gets the title from the <see cref="Assembly"/>.
        /// </summary>
        public string Title
        {
            get
            {
                return ((AssemblyTitleAttribute)Attribute.GetCustomAttribute(
                    this.assembly, typeof(AssemblyTitleAttribute), false))
                    .Title;
            } 
        }

        /// <summary>
        /// Gets the <see cref="Version"/> from the <see cref="Assembly"/>.
        /// </summary>
        public Version Version
        {
            get
            {
                return this.assembly.GetName().Version;
            }
        }

        /// <summary>
        /// Gets the code base path from the <see cref="Assembly"/>.
        /// </summary>
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
