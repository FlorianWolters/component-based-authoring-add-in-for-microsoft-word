//------------------------------------------------------------------------------
// <copyright file="FieldUpdater.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using FlorianWolters.Office.Word.Fields.UpdateStrategies;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="FieldUpdater"/> updates a collection of <see cref="Word.Fields"/> with a update
    /// <i>Strategy</i>.
    /// </summary>
    public class FieldUpdater
    {
        /// <summary>
        /// The update <i>Strategy</i> to use.
        /// </summary>
        private readonly IUpdateStrategy updateStrategy;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldUpdater"/> class with the specified update
        /// <i>Strategy</i>.
        /// </summary>
        /// <param name="updateStrategy">The update <i>Strategy</i> to use.</param>
        public FieldUpdater(IUpdateStrategy updateStrategy)
        {
            if (null == updateStrategy)
            {
                throw new ArgumentNullException("updateStrategy");
            }

            this.updateStrategy = updateStrategy;
        }

        /// <summary>
        /// Updates the specified collection of <see cref="Word.Field"/> objects with the current update
        /// <i>Strategy</i>.
        /// </summary>
        /// <param name="fields">The collection of <see cref="Word.Field"/> objects to update.</param>
        public void Update(IList<Word.Field> fields)
        {
            foreach (Word.Field field in fields)
            {
                this.updateStrategy.Update(field);
            }
        }

        /// <summary>
        /// Updates the specified collection of <see cref="Word.Field"/> objects with the current update
        /// <i>Strategy</i>.
        /// </summary>
        /// <param name="fields">The collection of <see cref="Word.Field"/> objects to update.</param>
        public void Update(Word.Fields fields)
        {
            foreach (Word.Field field in fields)
            {
                this.updateStrategy.Update(field);
            }
        }

        /// <summary>
        /// Updates the specified <see cref="Word.Field"/> with the current update <i>Strategy</i>.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> to update.</param>
        public void Update(Word.Field field)
        {
            this.updateStrategy.Update(field);
        }

        /// <summary>
        /// Updates all <see cref="Word.Field"/> objects in the specified <see cref="Word.Document"/>.
        /// <para>
        /// This method updates <b>all</b> fields in the specified document. This includes the content, and the headers
        /// and footers of all sections in the document.
        /// </para>
        /// </summary>
        /// <param name="document">
        /// The <see cref="Word.Document"/> whose <see cref="Word.Field"/> objects to update.
        /// </param>
        public void Update(Word.Document document)
        {
            this.Update(document.Sections);
            this.Update(document.Fields);
        }

        /// <summary>
        /// Updates all <see cref="Word.Field"/> objects in the specified <see cref="Word.Sections"/>.
        /// <para>
        /// This method updates <b>all</b> fields in the headers and footers of the specified sections.
        /// </para>
        /// </summary>
        /// <param name="sections">
        /// The <see cref="Word.Sections"/> whose <see cref="Word.Field"/> objects to update.
        /// </param>
        public void Update(Word.Sections sections)
        {
            foreach (Word.Section section in sections)
            {
                this.Update(section.Headers);
                this.Update(section.Footers);
            }
        }

        /// <summary>
        /// Updates all <see cref="Word.Field"/> objects in the specified <see cref="Word.HeadersFooters"/>.
        /// <para>
        /// This method updates <b>all</b> fields in the specified headers <b>or</b> footers.
        /// </para>
        /// </summary>
        /// <param name="headersFooters">
        /// The <see cref="Word.HeadersFooters"/> whose <see cref="Word.Field"/> objects to update.
        /// </param>
        public void Update(Word.HeadersFooters headersFooters)
        {
            foreach (Word.HeaderFooter headerFooter in headersFooters)
            {
                this.Update(headerFooter.Range.Fields);
            }
        }
    }
}
