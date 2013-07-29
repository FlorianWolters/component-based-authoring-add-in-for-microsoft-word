//------------------------------------------------------------------------------
// <copyright file="AssemblyInfoIntegrationTest.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------
namespace FlorianWolters.Reflection
{
    using System;
    using System.IO;
    using System.Reflection;
    using NUnit.Framework;

    [TestFixture]
    public class AssemblyInfoIntegrationTest
    {
        private AssemblyInfo assemblyInfo;

        [SetUp]
        public void SetUp()
        {
            this.assemblyInfo = new AssemblyInfo(Assembly.GetExecutingAssembly());
        }

        [Test]
        public void TestConstructorThrowsArgumentNullExceptionIfAssemblyIsNull()
        {
            Assert.That(
                () => new AssemblyInfo(null),
                Throws.Exception.TypeOf<ArgumentNullException>().With.Property("ParamName").EqualTo("assembly"));
        }

        [Test]
        public void TestCompany()
        {
            string expected = "Florian Wolters";
            string actual = this.assemblyInfo.Company;

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestCopyright()
        {
            string expected = "Copyright © Florian Wolters 2013";
            string actual = this.assemblyInfo.Copyright;

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestDescription()
        {
            string expected = "Automatic tests for the namespace FlorianWolters.Reflection";
            string actual = this.assemblyInfo.Description;

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestProduct()
        {
            string expected = "FlorianWolters.Reflection.Tests";
            string actual = this.assemblyInfo.Product;

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestTitle()
        {
            string expected = "FlorianWolters.Reflection.Tests";
            string actual = this.assemblyInfo.Title;

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestVersion()
        {
            Version expected = new Version("0.1.0.0");
            Version actual = this.assemblyInfo.Version;

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void TestCodeBasePath()
        {
            string actual = this.assemblyInfo.CodeBasePath;

            Assert.True(Directory.Exists(actual));
        }
    }
}
