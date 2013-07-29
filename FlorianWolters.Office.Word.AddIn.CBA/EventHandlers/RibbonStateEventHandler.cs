//------------------------------------------------------------------------------
// <copyright file="RibbonStateEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.EventHandlers
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using FlorianWolters.Office.Word.Fields.Switches;
    using Word = Microsoft.Office.Interop.Word;

    internal class RibbonStateEventHandler
        : IEventHandler,
        IDocumentChangeEventHandler,
        IDocumentBeforeSaveEventHandler,
        IWindowSelectionChangeEventHandler
    {
        private readonly Word.Application application;

        private readonly Ribbon ribbon;

        private readonly CustomDocumentPropertiesDropDown customDocumentPropertiesDropDown;

        public RibbonStateEventHandler(
            Word.Application application,
            Ribbon ribbon,
            CustomDocumentPropertiesDropDown customDocumentPropertiesDropDown)
        {
            this.application = application;
            this.ribbon = ribbon;
            this.customDocumentPropertiesDropDown = customDocumentPropertiesDropDown;
        }

        public void OnDocumentChange()
        {
            if (this.application.HasOpenDocuments())
            {
                this.ribbon.splitButtonFieldInsert.Enabled = true;
                this.ribbon.splitButtonFieldFormat.Enabled = true;
                this.ribbon.dropDownCustomDocumentProperties.Enabled = true;
                this.ribbon.checkBoxHideInternal.Enabled = true;
                this.customDocumentPropertiesDropDown.Update(this.ribbon.checkBoxHideInternal.Enabled);
                this.ribbon.buttonCreateCustomDocumentProperty.Enabled = true;
                this.UpdateDropDownFieldShading();
                this.UpdateToggleButtonShowFieldCodes();
                this.UpdateToggleButtonShowFieldShading();

                bool isSaved = this.application.ActiveDocument.IsSaved();
                this.ribbon.buttonInspect.Enabled = isSaved;
                this.ribbon.splitButtonInclude.Enabled = isSaved;
            }
            else
            {
                bool documentActive = false;
                this.customDocumentPropertiesDropDown.Clear();

                this.ribbon.buttonBindCustomXMLPart.Enabled = documentActive;
                this.ribbon.buttonInspect.Enabled = documentActive;
                this.ribbon.buttonOpenSourceFile.Enabled = documentActive;
                this.ribbon.buttonCreateCustomDocumentProperty.Enabled = documentActive;
                this.ribbon.buttonUpdateFromSource.Enabled = documentActive;
                this.ribbon.buttonUpdateToSource.Enabled = documentActive;
                this.ribbon.checkBoxHideInternal.Enabled = documentActive;
                this.ribbon.dropDownCustomDocumentProperties.Enabled = documentActive;
                this.ribbon.dropDownFieldShading.Enabled = documentActive;
                this.ribbon.splitButtonFieldFormat.Enabled = documentActive;
                this.ribbon.splitButtonInclude.Enabled = documentActive;
                this.ribbon.splitButtonFieldInsert.Enabled = documentActive;
                this.ribbon.toggleButtonShowFieldCode.Enabled = documentActive;
                this.ribbon.toggleButtonShowFieldCodes.Enabled = documentActive;
                this.ribbon.toggleButtonShowFieldShading.Enabled = documentActive;
            }
        }

        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            this.ribbon.buttonInspect.Enabled = true;
            this.ribbon.splitButtonInclude.Enabled = true;
        }

        public void OnWindowSelectionChange(Word.Selection selection)
        {
            IEnumerable<Word.Field> selectedFields = selection.SelectedFields();
            int selectedFieldsCount = selectedFields.Count();
            bool isSingleFieldSelected = 1 == selectedFieldsCount;

            if (isSingleFieldSelected)
            {
                Word.Field selectedField = selectedFields.ElementAt(0);
                FieldFunctionCode fieldFunctionCode = new FieldFunctionCode(selectedField.Code.Text);

                this.ribbon.buttonFieldUpdate.Enabled = !selectedField.Locked;
                this.ribbon.toggleButtonFieldLock.Checked = selectedField.Locked;
                this.ribbon.toggleButtonShowFieldCode.Checked = selectedField.ShowCodes;

                this.ribbon.toggleButtonFieldFormatAlphabetic.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Alphabetic);
                this.ribbon.toggleButtonFieldFormatArabic.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Arabic);
                this.ribbon.toggleButtonFieldFormatCaps.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Caps);
                this.ribbon.toggleButtonFieldFormatCardText.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.CardText);
                this.ribbon.toggleButtonFieldFormatDollarText.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.DollarText);
                this.ribbon.toggleButtonFieldFormatFirstCap.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.FirstCap);
                this.ribbon.toggleButtonFieldFormatHex.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Hex);
                this.ribbon.toggleButtonFieldFormatLower.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Lower);
                this.ribbon.toggleButtonFieldFormatMergeFormat.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.MergeFormat);
                this.ribbon.toggleButtonFieldFormatOrdinal.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Ordinal);
                this.ribbon.toggleButtonFieldFormatOrdText.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.OrdText);
                this.ribbon.toggleButtonFieldFormatRoman.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Roman);
                this.ribbon.toggleButtonFieldFormatUpper.Checked = fieldFunctionCode.ContainsFormatSwitch(FieldFormatSwitches.Upper);
            }

            this.ribbon.splitButtonFieldAction.Enabled = isSingleFieldSelected;
            this.ribbon.splitButtonFieldFormatCapitalization.Enabled = isSingleFieldSelected;
            this.ribbon.splitButtonFieldFormatNumber.Enabled = isSingleFieldSelected;
            this.UpdateSplitButtonInsertField(selection);
            this.UpdateButtonBindCustomXMLPart(selection);

            // TODO
            IEnumerable<Word.Field> selectedIncludeFields = selection.SelectedIncludeFields();
            bool includeFieldsAreSelected = selectedIncludeFields.Count() > 0;
            this.ribbon.buttonOpenSourceFile.Enabled = 1 == selectedIncludeFields.Count();
            this.ribbon.buttonUpdateFromSource.Enabled = includeFieldsAreSelected;
            this.ribbon.buttonUpdateToSource.Enabled = selection.SelectedIncludeTextFields().Count() > 0;
        }

        private void UpdateSplitButtonInsertField(Word.Selection selection)
        {
            this.ribbon.splitButtonFieldInsert.Enabled = selection.Start == selection.End;
        }

        private void UpdateToggleButtonShowFieldShading()
        {
            this.ribbon.toggleButtonShowFieldShading.Checked = this.application.ActiveDocument.FormFields.Shaded;
            this.ribbon.toggleButtonShowFieldShading.Enabled = true;
        }

        private void UpdateToggleButtonShowFieldCodes()
        {
            this.ribbon.toggleButtonShowFieldCodes.Checked = this.application.ActiveWindow.View.ShowFieldCodes;
            this.ribbon.toggleButtonShowFieldCodes.Enabled = true;
        }

        private void UpdateDropDownFieldShading()
        {
            Word.WdFieldShading wordFieldShading = this.application.ActiveWindow.View.FieldShading;

            int wordFieldShadingAsInt = (int)wordFieldShading;
            string wordFieldShadingAsString = wordFieldShadingAsInt.ToString();
            this.ribbon.dropDownFieldShading.SelectedItem = (from items in this.ribbon.dropDownFieldShading.Items
                                                             where items.Tag.Equals(wordFieldShadingAsString)
                                                             select items).First();
            this.ribbon.dropDownFieldShading.Enabled = true;
        }

        private void UpdateButtonBindCustomXMLPart(Word.Selection selection)
        {
            this.ribbon.buttonBindCustomXMLPart.Enabled = this.IsSingleContentControlSelected(selection);
        }

        private bool IsSingleContentControlSelected(Word.Selection selection)
        {
            return 1 == selection.Range.ContentControls.Count;
        }
    }
}
