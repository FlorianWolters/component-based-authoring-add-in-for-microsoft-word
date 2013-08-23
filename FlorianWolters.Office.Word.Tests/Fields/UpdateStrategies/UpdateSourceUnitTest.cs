//------------------------------------------------------------------------------
// <copyright file="UpdateSourceUnitTest.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using System;
    using System.Runtime.InteropServices;
    using Moq;
    using NUnit.Framework;
    using Word = Microsoft.Office.Interop.Word;

    [TestFixture]
    public class UpdateSourceUnitTest
    {
        [Test]
        public void TestUpdateThrowsArgumentException()
        {
            // Given
            Mock<Word.Field> mock = new Mock<Word.Field>();
            UpdateSource updateStrategy = new UpdateSource();

            // Expect
            mock.Setup(field => field.Type).Returns(Word.WdFieldType.wdFieldEmpty);

            // When and Then
            Assert.That(
                () => updateStrategy.Update(mock.Object),
                Throws.Exception.TypeOf<ArgumentException>().With.Property("ParamName").EqualTo("field"));
        }

        [Test]
        public void TestUpdateThrowsReadOnlyDocumentException()
        {
            // Given
            Mock<COMException> comExceptionMock = new Mock<COMException>();
            Mock<Word.Field> fieldMock = new Mock<Word.Field>();
            UpdateSource updateStrategy = new UpdateSource();

            // Expect
            comExceptionMock.Setup(ex => ex.ErrorCode).Returns(-2146823133);
            fieldMock.Setup(field => field.Type).Returns(Word.WdFieldType.wdFieldIncludeText);
            fieldMock.Setup(field => field.UpdateSource()).Throws(comExceptionMock.Object);

            // When and Then
            fieldMock.Verify(field => field.UpdateSource(), Times.Never());
            Assert.Throws<ReadOnlyDocumentException>(() => updateStrategy.Update(fieldMock.Object));
        }

        [Test]
        public void TestUpdate()
        {
            // Given
            Mock<Word.Field> mock = new Mock<Word.Field>();
            UpdateSource updateStrategy = new UpdateSource();

            // Expect
            mock.Setup(field => field.Type).Returns(Word.WdFieldType.wdFieldIncludeText);

            // When
            updateStrategy.Update(mock.Object);

            // Then
            mock.Verify(field => field.UpdateSource(), Times.Once());
        }
    }
}
