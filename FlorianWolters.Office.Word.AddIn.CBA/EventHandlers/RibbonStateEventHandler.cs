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

    /// <summary>
    /// The class <see cref="RibbonStateEventHandler"/> implements <i>Event Handler</i> methods which modify the state
    /// of the user interface of the <see cref="Ribbon"/>.
    /// </summary>
    internal class RibbonStateEventHandler
        : IEventHandler, IDocumentChangeEventHandler, IDocumentBeforeSaveEventHandler, IWindowSelectionChangeEventHandler
    {
        /// <summary>
        /// The <see cref="Word.Application"/> to react to.
        /// </summary>
        private readonly Word.Application application;

        /// <summary>
        /// The <see cref="Ribbon"/> whose user interface state to modify.
        /// </summary>
        private readonly Ribbon ribbon;

        /// <summary>
        /// Allows to modify the dropdown control with the custom document properties.
        /// </summary>
        private readonly CustomDocumentPropertiesDropDown customDocumentPropertiesDropDown;

        /// <summary>
        /// Initializes a new instance of the <see cref="RibbonStateEventHandler"/> class.
        /// </summary>
        /// <param name="application">The <see cref="Word.Application"/> to react to.</param>
        /// <param name="ribbon">The <see cref="Ribbon"/> whose user interface state to modify.</param>
        /// <param name="customDocumentPropertiesDropDown">Allows to modify the dropdown control with the custom document properties.</param>
        public RibbonStateEventHandler(
            Word.Application application,
            Ribbon ribbon,
            CustomDocumentPropertiesDropDown customDocumentPropertiesDropDown)
        {
            this.application = application;
            this.ribbon = ribbon;
            this.customDocumentPropertiesDropDown = customDocumentPropertiesDropDown;
        }

        /// <summary>
        /// Handles the event which occurs when a new <see cref="Microsoft.Office.Interop.Word.Document"/> is created,
        /// when an existing <see cref="Microsoft.Office.Interop.Word.Document"/> is opened, or when another <see
        /// cref="Microsoft.Office.Interop.Word.Document"/> is made the active <see
        /// cref="Microsoft.Office.Interop.Word.Document"/>. 
        /// </summary>
        public void OnDocumentChange()
        {
            if (this.application.HasOpenDocuments())
            {
                bool activeDocumentSaved = this.application.ActiveDocument.IsSaved();

                this.ribbon.splitButtonInclude.Enabled = activeDocumentSaved;
                this.ribbon.buttonBindCustomXMLPart.Enabled = activeDocumentSaved;

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

        /// <summary>
        /// Handles the event which occurs before any <see cref="Word.Document"/> is saved.
        /// </summary>
        /// <param name="document">The <see cref="Word.Document"/> that's being saved.</param>
        /// <param name="saveAsUI">
        /// <c>true</c> if the <b>Save As</b> dialog box is displayed, whether to save a new <see
        /// cref="Word.Document"/>, in response to the <b>Save</b> command; or in response to the <b>Save As</b>
        /// command; or in response to the <b>SaveAs</b> or <b>SaveAs2</b> method.
        /// </param>
        /// <param name="cancel">
        /// <c>false</c> when the event occurs. If the event procedure sets this argument to <c>true</c>, the <see
        /// cref="Word.Document"/> is not saved when the procedure is finished.
        /// </param>
        public void OnDocumentBeforeSave(Word.Document document, ref bool saveAsUI, ref bool cancel)
        {
            this.ribbon.splitButtonInclude.Enabled = true;
        }

        /// <summary>
        /// Handles the event which occurs when the <see cref="Word.Selection"/> changes in the active <see
        /// cref="Word.Document"/> <see cref="Word.Window"/>.
        /// </summary>
        /// <param name="selection">
        /// The text selected. If no text is selected, this parameter is either <c>null</c> or contains the first
        /// character to the right of the insertion point.
        /// </param>
        public void OnWindowSelectionChange(Word.Selection selection)
        {
            this.ribbon.splitButtonFieldInsert.Enabled = selection.Start == selection.End;
            this.ribbon.buttonBindCustomXMLPart.Enabled = 0 == selection.Range.ContentControls.Count;

            if (selection.Range.ContentControls.Count > 0)
            {
                return;
            }

            // TODO The extension method AllFields is too slow. Find a better solution.
            // Meanwhile we stick with selection.Range.Field.
            ////IList<Word.Field> selectedFields = selection.AllFields().ToList();
            IList<Word.Field> selectedFields = new List<Word.Field>(selection.Range.Fields.Cast<Word.Field>());
            int selectedFieldCount = selectedFields.Count();

            bool fieldsSelected = 0 < selectedFieldCount;
            bool singleFieldSelected = 1 == selectedFieldCount;
            bool oneOrMoreFieldsLocked = false;
            bool oneOrMoreIncludeTextFields = false;
            bool oneOrMoreIncludePictureFields = false;
            bool oneOrMoreIncludeFields = false;

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

                int includeTextFieldCount = (from f in selectedFields
                                             where f.Type == Word.WdFieldType.wdFieldIncludeText
                                             select f).Count();
                int includePictureFieldCount = (from f in selectedFields
                                                where f.Type == Word.WdFieldType.wdFieldIncludePicture
                                                select f).Count();
                oneOrMoreIncludeTextFields = 0 < includeTextFieldCount;
                oneOrMoreIncludePictureFields = 0 < includePictureFieldCount;
                oneOrMoreIncludeFields = oneOrMoreIncludeTextFields || oneOrMoreIncludePictureFields;

                this.ribbon.buttonFieldUpdate.Enabled = !oneOrMoreFieldsLocked;

                this.ribbon.toggleButtonFieldLock.Checked = oneOrMoreFieldsLocked;
                
                // Because of the problematic of duplicated InsertPicture fields a IncludePicture field cannot be
                // locked. If a field is locked, it can't be updated. If a IncludePicture field can't be updated it is
                // duplicated in the OOXML.
                this.ribbon.toggleButtonFieldLock.Enabled = (singleFieldSelected
                    || lockedFieldCount == 0
                    || lockedFieldCount == selectedFieldCount)
                    && !oneOrMoreIncludePictureFields;

                this.ribbon.toggleButtonFieldShowCode.Checked = oneOrMoreFieldsShowCodes;
                this.ribbon.toggleButtonFieldShowCode.Enabled = singleFieldSelected
                    || showCodesFieldCount == 0
                    || showCodesFieldCount == selectedFieldCount;

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

            this.ribbon.menuFieldAction.Enabled = fieldsSelected;
            this.ribbon.menuFieldFormat.Enabled = fieldsSelected && !oneOrMoreIncludeFields;
            this.ribbon.buttonUpdateFromSource.Enabled = oneOrMoreIncludeFields && !oneOrMoreFieldsLocked;
            this.ribbon.buttonOpenSourceFile.Enabled = oneOrMoreIncludeFields;
            this.ribbon.buttonUpdateToSource.Enabled = oneOrMoreIncludeTextFields;
            this.ribbon.buttonCompare.Enabled = oneOrMoreIncludeTextFields;
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
