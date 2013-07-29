//------------------------------------------------------------------------------
// <copyright file="FieldUpdater.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.Fields
{
    using System.Collections.Generic;
    using FlorianWolters.Office.Word.Fields.UpdateStrategies;
    using Word = Microsoft.Office.Interop.Word;

    public class FieldUpdater
    {
        private readonly IEnumerable<Word.Field> fields;

        public FieldUpdater(
            IEnumerable<Word.Field> fields,
            IUpdateStrategy strategy)
        {
            this.fields = fields;
            this.Strategy = strategy;
        }

        public FieldUpdater(
            Word.Field field,
            IUpdateStrategy strategy)
            : this(new[] { field }, strategy)
        {
        }

        public IUpdateStrategy Strategy { get; set; }

        public void Update()
        {
            foreach (Word.Field field in this.fields)
            {
                this.Strategy.Update(field);
            }
        }
    }
}
