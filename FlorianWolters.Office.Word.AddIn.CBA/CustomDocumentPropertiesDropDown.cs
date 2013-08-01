//------------------------------------------------------------------------------
// <copyright file="CustomDocumentPropertiesDropDown.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA
{
    using System.Collections.Generic;
    using FlorianWolters.Office.Word.DocumentProperties;
    using Microsoft.Office.Tools.Ribbon;
    using Office = Microsoft.Office.Core;
    using Word = Microsoft.Office.Interop.Word;

    internal class CustomDocumentPropertiesDropDown
    {
        private readonly Word.Application application;
        private readonly RibbonFactory ribbonFactory;
        private readonly RibbonDropDown dropDown;
        private readonly CustomDocumentPropertyReader customDocumentPropertyReader;

        public CustomDocumentPropertiesDropDown(
            Word.Application application,
            RibbonFactory ribbonFactory,
            RibbonDropDown dropDown,
            CustomDocumentPropertyReader customDocumentPropertyReader)
        {
            // TODO Validate parameters.
            this.application = application;
            this.ribbonFactory = ribbonFactory;
            this.dropDown = dropDown;
            this.customDocumentPropertyReader = customDocumentPropertyReader;
        }

        public void ResetSelection()
        {
            this.dropDown.SelectedItemIndex = 0;
        }

        public void Clear()
        {
            this.dropDown.Items.Clear();
        }

        public void Update(bool hideInternal = true)
        {
            this.customDocumentPropertyReader.Load(this.application.ActiveDocument);
            IEnumerable<Office.DocumentProperty> customDocumentProperties = hideInternal
                ? this.customDocumentPropertyReader.FindAllExceptInternal()
                : customDocumentProperties = this.customDocumentPropertyReader.GetAll();

            this.Clear();
            this.AddItemToCustomPropertiesDropDownList("Select a property to insert");

            foreach (Office.DocumentProperty property in customDocumentProperties)
            {
                this.AddItemToCustomPropertiesDropDownList(property);
            }
        }

        private void AddItemToCustomPropertiesDropDownList(Office.DocumentProperty customProperty)
        {
            string label = customProperty.Name;
            string superTip = "Select to insert the value ('"
                + customProperty.Value
                + "') of the custom document property with the name '"
                + customProperty.Name
                + "' at the current position into the document.";
            this.AddItemToCustomPropertiesDropDownList(label, superTip);
        }

        private void AddItemToCustomPropertiesDropDownList(string label, string superTip = null)
        {
            RibbonDropDownItem item = this.ribbonFactory.CreateRibbonDropDownItem();
            item.Label = label;
            item.SuperTip = superTip;
            this.dropDown.Items.Add(item);
        }
    }
}
