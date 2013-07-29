//------------------------------------------------------------------------------
// <copyright file="FieldFunctionCodeUnitTest.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System;
    using NUnit.Framework;

    [TestFixture]
    public class FieldFunctionCodeUnitTest
    {
        private static readonly object[] FormatCases =
        {
            new object[] { new FieldFunctionCode("AUTHOR"), string.Empty },
            new object[] { new FieldFunctionCode(@"AUTHOR \*Lower"), @"\*LOWER" },
            new object[] { new FieldFunctionCode(@"AUTHOR \*MERGEFORMAT \*Lower"), @"\*MERGEFORMAT \*LOWER" },
            new object[] { new FieldFunctionCode(@" AUTHOR  \* Lower "), @"\*LOWER" },
            new object[]
            {
                new FieldFunctionCode(
                    @" INCLUDETEXT  ""C:\\Path with multiple   whitespaces\\Document.docx"" \* Upper "),
                    @"\*UPPER"
            }
        };

        private static readonly object[] FunctionCases =
        {
            new object[] { new FieldFunctionCode("AUTHOR"), "AUTHOR" },
            new object[] { new FieldFunctionCode(@"AUTHOR \*Lower"), "AUTHOR" },
            new object[] { new FieldFunctionCode(@"AUTHOR \*MERGEFORMAT \*Lower"), "AUTHOR" },
            new object[] { new FieldFunctionCode(@" AUTHOR  \* Lower "), "AUTHOR" },
            new object[]
            {
                new FieldFunctionCode(
                    @" INCLUDETEXT  ""C:\\Path with multiple   whitespaces\\Document.docx"" \* Upper "),
                    @"INCLUDETEXT  ""C:\\Path with multiple   whitespaces\\Document.docx"""
            }
        };

        [Test, TestCaseSource("FormatCases")]
        public void TestFormat(FieldFunctionCode fieldFunctionCode, string expected)
        {
            string actual = fieldFunctionCode.Format;

            Assert.AreEqual(expected, actual);
        }

        [Test, TestCaseSource("FunctionCases")]
        public void TestFunction(FieldFunctionCode fieldFunctionCode, string expected)
        {
            string actual = fieldFunctionCode.Function;

            Assert.AreEqual(expected, actual);
        }
    }
}
