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

    /// <summary>
    /// The class <see cref="CustomDocumentPropertiesDropDown"/> decorates a <see cref="RibbonDropDown"/> to interact
    /// with custom document properties of a  Microsoft Word document.
    /// </summary>
    internal class CustomDocumentPropertiesDropDown
    {
        /// <summary>
        /// The <see cref="Word.Application"/> to retrieve the active <see cref="Word.Document"/>.
        /// </summary>
        private readonly Word.Application application;

        /// <summary>
        /// The <see cref="RibbonFactory"/> to create <see cref="RibbonDropDownItem"/> objects.
        /// </summary>
        private readonly RibbonFactory ribbonFactory;

        /// <summary>
        /// The <see cref="RibbonDropDown"/> to decorate.
        /// </summary>
        private readonly RibbonDropDown dropDown;

        /// <summary>
        /// The <see cref="CustomDocumentPropertyReader"/> to read the custom document properties from the active <see
        /// cref="Word.Document"/>.
        /// </summary>
        private readonly CustomDocumentPropertyReader customDocumentPropertyReader;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomDocumentPropertiesDropDown"/> class.
        /// </summary>
        /// <param name="application">
        /// The <see cref="Word.Application"/> to retrieve the active <see cref="Word.Document"/>.
        /// </param>
        /// <param name="ribbonFactory">
        /// The <see cref="RibbonFactory"/> to create <see cref="RibbonDropDownItem"/> objects.
        /// </param>
        /// <param name="dropDown">The <see cref="RibbonDropDown"/> to decorate.</param>
        /// <param name="customDocumentPropertyReader">
        /// The <see cref="CustomDocumentPropertyReader"/> to read the custom document properties from the active <see
        /// cref="Word.Document"/>.
        /// </param>
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

        /// <summary>
        /// Resets the current selection in the <see cref="RibbonDropDown"/>.
        /// </summary>
        public void ResetSelection()
        {
            this.dropDown.SelectedItemIndex = 0;
        }

        /// <summary>
        /// Removes all items from the <see cref="RibbonDropDown"/>.
        /// </summary>
        public void Clear()
        {
            this.dropDown.Items.Clear();
        }

        /// <summary>
        /// Updates the <see cref="RibbonDropDown"/>.
        /// </summary>
        /// <param name="hideInternal">Whether to hide <i>internal</i> custom document properties.</param>
        /// <remarks><i>Internal</i> custom document properties start with an underscore character <c>_</c>.</remarks>
        public void Update(bool hideInternal = true)
        {
            this.customDocumentPropertyReader.Load(this.application.ActiveDocument);
            IEnumerable<Office.DocumentProperty> customDocumentProperties = hideInternal
                ? this.customDocumentPropertyReader.FindAllExceptInternal()
                : this.customDocumentPropertyReader.GetAll();

            this.Clear();
            this.AddItemToCustomPropertiesDropDownList("Select a property to insert");

            foreach (Office.DocumentProperty property in customDocumentProperties)
            {
                this.AddItemToCustomPropertiesDropDownList(property);
            }
        }

        /// <summary>
        /// Creates a new <see cref="RibbonDropDownItem"/> from the specified <see cref="Office.DocumentProperty"/> and
        /// adds it to the <see cref="RibbonDropDown"/>.
        /// </summary>
        /// <param name="customProperty">The <see cref="Office.DocumentProperty"/> to add.</param>
        private void AddItemToCustomPropertiesDropDownList(Office.DocumentProperty customProperty)
        {
            string label = customProperty.Name;
            string superTip = "Select to insert the value ('"
                + customProperty.Value
                + "') of the custom document property with the name '"
                + customProperty.Name
                + "' at the current position into the document.";

            RibbonDropDownItem item = this.ribbonFactory.CreateRibbonDropDownItem();
            item.Label = label;
            item.SuperTip = superTip;
            this.dropDown.Items.Add(item);
        }

        /// <summary>
        // Creates a new <see cref="RibbonDropDownItem"/> with the specified label and screen tip.
        /// </summary>
        /// <param name="label">The label of the <see cref="RibbonDropDownItem"/>.</param>
        /// <param name="screenTip">The screen tip of the <see cref="RibbonDropDownItem"/>.</param>
        private void AddItemToCustomPropertiesDropDownList(string label, string screenTip = null)
        {
            RibbonDropDownItem item = this.ribbonFactory.CreateRibbonDropDownItem();
            item.Label = label;
            item.ScreenTip = screenTip;
            this.dropDown.Items.Add(item);
        }
    }
}
