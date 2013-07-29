//------------------------------------------------------------------------------
// <copyright file="PathUtilsIntegrationTest.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.IO.Tests
{
    using System;
    using System.IO;
    using System.Reflection;
    using NUnit.Framework;

    [TestFixture]
    public class PathUtilsIntegrationTest
    {
        private const string NonExistentDirectory = @"C:\System Volume Information\foo";

        private static readonly object[] ArgumentNullExceptionCases =
        {
            new object[] { null },
            new object[] { string.Empty }
        };

        private static readonly object[] FormatCases =
        {
            new object[] { ".", ".", "." },
            new object[] { "C:", "C:", "." },
            new object[] { "C:\\", "C:", "." },
            new object[] { "C:", "C:\\", "." },
            new object[] { "C:", "C:\\Windows", ".\\Windows" },
            new object[] { "C:\\Windows", "C:\\", ".." },
            new object[] { "C:\\Windows\\System32\\Drivers", "C:\\System Volume Information", "..\\..\\..\\System Volume Information" },
            new object[] { "C:\\System Volume Information", "C:\\Windows\\System32\\Drivers", "..\\Windows\\System32\\Drivers" },
            new object[] { "C:\\bootmgr", "C:\\Windows\\System32", ".\\Windows\\System32" }
        };

        [Test, TestCaseSource("FormatCases")]
        public void TestGetRelativePath(string fromPath, string toPath, string expected)
        {
            string actual = PathUtils.GetRelativePath(fromPath, toPath);

            Assert.AreEqual(expected, actual);
        }

        [Test, TestCaseSource("ArgumentNullExceptionCases")]
        public void TestGetRelativePathThrowsArgumentNullExceptionIfFromPathIsNull(string fromPath)
        {
            Assert.That(
                () => PathUtils.GetRelativePath(fromPath, "."),
                Throws.Exception.TypeOf<ArgumentNullException>().With.Property("ParamName").EqualTo("fromPath"));
        }

        [Test, TestCaseSource("ArgumentNullExceptionCases")]
        public void TestGetRelativePathThrowsArgumentNullExceptionIfToPathIsNull(string toPath)
        {
            Assert.That(
                () => PathUtils.GetRelativePath(".", toPath),
                Throws.Exception.TypeOf<ArgumentNullException>().With.Property("ParamName").EqualTo("toPath"));
        }

        [Test]
        public void TestGetRelativePathThrowsArgumentNullExceptionIfFromPathDoesNotExist()
        {
            Assert.That(
                () => PathUtils.GetRelativePath(NonExistentDirectory, "."),
                Throws.Exception.TypeOf<FileNotFoundException>());
        }

        [Test]
        public void TestGetRelativePathThrowsArgumentNullExceptionIfFromPathToNotExist()
        {
            Assert.That(
                () => PathUtils.GetRelativePath(".", NonExistentDirectory),
                Throws.Exception.TypeOf<FileNotFoundException>());
        }

        [Test]
        public void TestGetRelativePathThrowsArgumentExceptionIfPathsDoNotHaveCommonPathPrefix()
        {
            Assert.That(
                () => PathUtils.GetRelativePath("C:", "D:"),
                Throws.Exception.TypeOf<ArgumentException>().With.Property("Message").EqualTo("The paths must have a common prefix."));
        }
    }
}
