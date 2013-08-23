//------------------------------------------------------------------------------
// <copyright file="UpdateTargetUnitTest.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields.UpdateStrategies
{
    using Moq;
    using NUnit.Framework;
    using Word = Microsoft.Office.Interop.Word;

    [TestFixture]
    public class UpdateTargetUnitTest
    {
        [Test]
        public void TestUpdate()
        {
            // Given
            Mock<Word.Field> mock = new Mock<Word.Field>();
            UpdateTarget updateTarget = new UpdateTarget();

            // Expect
            mock.Setup(field => field.Update())
                .Returns(true);

            // When
            updateTarget.Update(mock.Object);

            // Then
            mock.Verify(field => field.Update(), Times.Once());
        }
    }
}
