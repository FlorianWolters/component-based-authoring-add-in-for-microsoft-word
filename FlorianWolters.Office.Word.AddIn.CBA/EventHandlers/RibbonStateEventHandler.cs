//------------------------------------------------------------------------------
// <copyright file="RibbonStateEventHandler.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.EventHandlers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
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
                bool activeDocumentSaved = this.application.ActiveDocument.IsSaved();

                this.ribbon.splitButtonInclude.Enabled = activeDocumentSaved;
                this.ribbon.buttonCompare.Enabled = activeDocumentSaved;

                this.customDocumentPropertiesDropDown.Update(this.ribbon.checkBoxHideInternal.Checked);
                this.ribbon.dropDownCustomDocumentProperties.Enabled = true;
                this.ribbon.checkBoxHideInternal.Enabled = true;
                this.ribbon.buttonCreateCustomDocumentProperty.Enabled = true;
                
                this.ribbon.splitButtonFieldInsert.Enabled = true;

                this.ribbon.toggleButtonFormFieldShading.Enabled = true;
                this.ribbon.toggleButtonFieldCodes.Enabled = true;
                this.ribbon.toggleButtonFieldCodes.Checked = this.application.ActiveWindow.View.ShowFieldCodes;

                try
                {
                    // TODO Find a way to avoid the throwing of the exception,
                    // e.g. by implementing a detection of the comparison view.
                    this.ribbon.toggleButtonFormFieldShading.Checked = this.application.ActiveDocument.FormFields.Shaded;
                    this.UpdateDropDownFieldShading();
                }
                catch (COMException)
                {
                    // We simply "swallow" a possible COMException (required if
                    // two Word documents are compared).
                    this.ribbon.toggleButtonFormFieldShading.Enabled = false;
                    this.ribbon.dropDownFieldShading.Enabled = false;
                }
            }
            else
            {
                bool documentActive = false;

                this.ribbon.splitButtonInclude.Enabled = documentActive;
                this.ribbon.buttonUpdateFromSource.Enabled = documentActive;
                this.ribbon.buttonOpenSourceFile.Enabled = documentActive;
                this.ribbon.buttonUpdateToSource.Enabled = documentActive;
                this.ribbon.buttonCompare.Enabled = documentActive;

                this.ribbon.buttonBindCustomXMLPart.Enabled = documentActive;

                this.ribbon.splitButtonFieldInsert.Enabled = documentActive;
                this.ribbon.menuFieldFormat.Enabled = documentActive;
                this.ribbon.menuFieldAction.Enabled = documentActive;

                this.customDocumentPropertiesDropDown.Clear();
                this.ribbon.dropDownCustomDocumentProperties.Enabled = documentActive;
                this.ribbon.checkBoxHideInternal.Enabled = documentActive;
                this.ribbon.buttonCreateCustomDocumentProperty.Enabled = documentActive;

                this.ribbon.dropDownFieldShading.Enabled = documentActive;
                this.ribbon.toggleButtonFormFieldShading.Enabled = documentActive;
                this.ribbon.toggleButtonFieldCodes.Enabled = documentActive;
            }
        }

        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            this.ribbon.buttonCompare.Enabled = true;
            this.ribbon.splitButtonInclude.Enabled = true;
        }

        public void OnWindowSelectionChange(Word.Selection selection)
        {
            this.ribbon.splitButtonFieldInsert.Enabled = selection.Start == selection.End;
            this.ribbon.buttonBindCustomXMLPart.Enabled = 0 == selection.Range.ContentControls.Count;

            IEnumerable<Word.Field> selectedFields = selection.AllFields();
            int selectedFieldCount = selectedFields.Count();

            bool fieldsSelected = 0 < selectedFieldCount;
            bool singleFieldSelected = 1 == selectedFieldCount;
            bool oneOrMoreFieldsLocked = false;
            bool oneOrMoreIncludeFields = false;

            this.ribbon.menuFieldAction.Enabled = fieldsSelected;
            this.ribbon.menuFieldFormat.Enabled = fieldsSelected;

            if (fieldsSelected)
            {
                int showCodesFieldCount = (from f in selectedFields
                                        where f.ShowCodes == true
                                        select f).Count();
                bool oneOrMoreFieldsShowCodes = 0 < showCodesFieldCount;

                int lockedFieldCount = (from f in selectedFields
                                        where f.Locked == true
                                        select f).Count();
                oneOrMoreFieldsLocked = 0 < lockedFieldCount;

                this.ribbon.buttonFieldUpdate.Enabled = !oneOrMoreFieldsLocked;
                
                this.ribbon.toggleButtonFieldLock.Checked = oneOrMoreFieldsLocked;
                this.ribbon.toggleButtonFieldLock.Enabled = singleFieldSelected
                    || lockedFieldCount == 0
                    || lockedFieldCount == selectedFieldCount;
                
                this.ribbon.toggleButtonFieldShowCode.Checked = oneOrMoreFieldsShowCodes;
                this.ribbon.toggleButtonFieldShowCode.Enabled = singleFieldSelected
                    || showCodesFieldCount == 0
                    || showCodesFieldCount == selectedFieldCount;

                int includeFieldCount = (from f in selectedFields
                                         where f.Type == Word.WdFieldType.wdFieldIncludeText
                                            || f.Type == Word.WdFieldType.wdFieldIncludePicture
                                            || f.Type == Word.WdFieldType.wdFieldInclude
                                         select f).Count();
                oneOrMoreIncludeFields = 0 < includeFieldCount;

                if (singleFieldSelected)
                {
                    Word.Field selectedField = selectedFields.ElementAt(0);
                    FieldFunctionCode fieldFunctionCode = new FieldFunctionCode(selectedField.Code.Text);

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
            }

            this.ribbon.buttonUpdateFromSource.Enabled = oneOrMoreIncludeFields && !oneOrMoreFieldsLocked;
            this.ribbon.buttonOpenSourceFile.Enabled = oneOrMoreIncludeFields;
            this.ribbon.buttonUpdateToSource.Enabled = oneOrMoreIncludeFields;
            this.ribbon.buttonCompare.Enabled = oneOrMoreIncludeFields;
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
    }
}
