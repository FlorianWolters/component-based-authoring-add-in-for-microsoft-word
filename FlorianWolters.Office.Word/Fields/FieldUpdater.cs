//------------------------------------------------------------------------------
// <copyright file="FieldUpdater.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
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
        /// The collection of <see cref="Word.Field"/> objects to update.
        /// </summary>
        private readonly IList<Word.Field> fields;

        /// <summary>
        /// The update <i>Strategy</i> to use.
        /// </summary>
        private readonly IUpdateStrategy updateStrategy;

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldUpdater"/> class.
        /// </summary>
        /// <param name="fields">The collection of <see cref="Word.Field"/> objects to update.</param>
        /// <param name="updateStrategy">The update <i>Strategy</i> to use.</param>
        public FieldUpdater(Word.Fields fields, IUpdateStrategy updateStrategy)
            : this(new List<Word.Field>(fields.Cast<Word.Field>()), updateStrategy)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldUpdater"/> class.
        /// </summary>
        /// <param name="fields">The collection of <see cref="Word.Field"/> objects to update.</param>
        /// <param name="updateStrategy">The update <i>Strategy</i> to use.</param>
        public FieldUpdater(IList<Word.Field> fields, IUpdateStrategy updateStrategy)
        {
            this.fields = fields;
            this.updateStrategy = updateStrategy;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FieldUpdater"/> class.
        /// </summary>
        /// <param name="field">The <see cref="Word.Field"/> object to update.</param>
        /// <param name="updateStrategy">The update <i>Strategy</i> to use.</param>
        public FieldUpdater(Word.Field field, IUpdateStrategy updateStrategy)
            : this(new[] { field }, updateStrategy)
        {
        }

        /// <summary>
        /// Updates the <see cref="Word.Field"/> object(s) with the update <i>Strategy</i>.
        /// </summary>
        public void Update()
        {
            foreach (Word.Field field in this.fields)
            {
                this.updateStrategy.Update(field);
            }
        }
    }
}
