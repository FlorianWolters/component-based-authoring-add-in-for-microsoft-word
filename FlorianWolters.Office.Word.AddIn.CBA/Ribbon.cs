//------------------------------------------------------------------------------
// <copyright file="Ribbon.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.CustomXML;
    using FlorianWolters.Office.Word.AddIn.CBA.EventHandlers;
    using FlorianWolters.Office.Word.AddIn.CBA.Factories;
    using FlorianWolters.Office.Word.AddIn.CBA.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.Dialogs;
    using FlorianWolters.Office.Word.DocumentProperties;
    using FlorianWolters.Office.Word.Event;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Event.ExceptionHandlers;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using FlorianWolters.Office.Word.Fields.Switches;
    using FlorianWolters.Office.Word.Fields.UpdateStrategies;
    using FlorianWolters.Reflection;
    using FlorianWolters.Windows.Forms;
    using Microsoft.Office.Tools.Ribbon;
    using Office = Microsoft.Office.Core;
    using VB = Microsoft.VisualBasic;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// The class <see cref="Ribbon"/> contains the presentation logic (the
    /// <i>Ribbon</i> and delegates to the business logic of the <i>Microsoft
    /// Word</i> Application-Level Add-In.
    /// </summary>
    internal partial class Ribbon
    {
        /// <summary>
        /// The file name of the ReadMe file.
        /// </summary>
        private const string ReadMeFileName = "README.md";

        /// <summary>
        /// The <see cref="Word.Application"/>, this <see cref="Ribbon"/> is
        /// running in.
        /// </summary>
        private Word.Application application;

        /// <summary>
        /// Gets or sets the windows of the <i>Microsoft Word</i> Application,
        /// this <see cref="Ribbon"/> is running in.
        /// </summary>
        private IWin32Window ApplicationWindow { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="AboutForm"/>.
        /// </summary>
        private AboutForm AboutDialog { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="MarkdownForm"/>.
        /// </summary>
        private MarkdownForm HelpDialog { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="ConfigurationForm"/>.
        /// </summary>
        private ConfigurationForm ConfigurationDialog { get; set; }

        /// <summary>
        /// Gets or sets the <see cref="CustomXMLPartsForm"/>.
        /// </summary>
        private CustomXMLPartsForm CustomXMLPartsDialog { get; set; }

        private CustomDocumentPropertiesDropDown CustomDocumentPropertiesDropDown { get; set; }

        private FieldFactory FieldFactory { get; set; }

        /// <summary>
        /// Occurs when this <see cref="Ribbon"/> is loaded into the Microsoft
        /// Office application. 
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnLoad(object sender, RibbonUIEventArgs e)
        {
            // This is the suggested way to access the Add-In object and the
            // Microsoft Word application object.
            // The global Add-In object is first available when the event
            // handler for the "Load" event of the "Ribbon" is invoked.
            
            // TODO Remove member variable if unused.
            this.application = Globals.ThisAddIn.Application;

            // TODO Validate configuration options.
            Settings settings = Settings.Default;

            AssemblyInfo assemblyInfo = new AssemblyInfo(Assembly.GetExecutingAssembly());
            this.logger.Info("Loaded " + Settings.Default.ApplicationName + " v" + assemblyInfo.Version.ToString() + ".");

            CustomDocumentPropertyReader customDocumentPropertyReader = new CustomDocumentPropertyReader();
            this.FieldFactory = new FieldFactory(this.application, customDocumentPropertyReader);
            this.CustomDocumentPropertiesDropDown = new CustomDocumentPropertiesDropDown(
                this.application,
                this.Factory,
                this.dropDownCustomDocumentProperties,
                customDocumentPropertyReader);

            this.InitializeForms(settings, assemblyInfo);
            this.RegisterEventHandler(settings);
        }

        /// <summary>
        /// Occurs when this <see cref="Ribbon"/> instance is closing.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">An object that contains no event data.</param>
        private void OnClose(object sender, EventArgs e)
        {
            // NOOP
        }

        /// <summary>
        /// Initializes the forms used by this <see cref="Ribbon"/>.
        /// </summary>
        /// <param name="assemblyInfo"></param>
        private void InitializeForms(Settings settings, AssemblyInfo assemblyInfo)
        {
            this.ConfigurationDialog = new ConfigurationForm(settings);
            this.AboutDialog = new AboutForm(assemblyInfo, settings);
            this.InitializeHelpDialog(assemblyInfo);
            this.CustomXMLPartsDialog = new CustomXMLPartsForm();
        }

        private void InitializeHelpDialog(AssemblyInfo assemblyInfo)
        {
            string readMeFilePath = assemblyInfo.CodeBasePath + Path.DirectorySeparatorChar + ReadMeFileName;

            try
            {
                this.HelpDialog = new MarkdownForm(readMeFilePath);
                this.HelpDialog.ChangeTitle("Help");
            }
            catch (FileNotFoundException)
            {
                MessageBoxes.ShowMessageBoxHelpFieldDoesNotExist(readMeFilePath);
                this.buttonShowHelpForm.Enabled = false;
            }
        }

        // TODO How to configure this?

        /// <summary>
        /// Registers all <i>Event Handlers</i> for this <see cref="Ribbon"/>.
        /// </summary>
        private void RegisterEventHandler(Settings settings)
        {
            IExceptionHandler eventExceptionHandler = new LoggerExceptionHandler(this.logger);

            // ATTENTION: Since we always inject the Word.Application into the commands, we can always access the current state of the Microsoft Word application.
            // If we would work with Word.Document instead, we would always have to make sure that the reference to the document is up-to-date.
            ApplicationEventHandler applicationEventHandler = new ApplicationEventHandler(this.application);

            // TODO Improve registration of the event handlers in dependency of the settings.

            if (settings.ActivateUpdateStylesOnOpen)
            {
                ActivateUpdateStylesOnOpenFactory.Instance.RegisterEventHandler(eventExceptionHandler, applicationEventHandler);
            }

            if (settings.RefreshCustomXMLParts)
            {
                RefreshCustomXMLPartsFactory.Instance.RegisterEventHandler(eventExceptionHandler, applicationEventHandler);
            }

            if (settings.UpdateAttachedTemplate)
            {
                UpdateAttachedTemplateFactory.Instance.RegisterEventHandler(eventExceptionHandler, applicationEventHandler);

            }

            if (settings.UpdateFields)
            {
                UpdateFieldsFactory.Instance.RegisterEventHandler(eventExceptionHandler, applicationEventHandler);
            }


            if (settings.WriteCustomDocumentProperties)
            {
                WriteCustomDocumentPropertiesFactory.Instance.RegisterEventHandler(eventExceptionHandler, applicationEventHandler);
            }

            // The RibbonStateEventHandler ensures that the state of the UI of this Ribbon is correctly set.
            // TODO Refactor.
            IEventHandler eventHandler = new RibbonStateEventHandler(
                this.application,
                this,
                this.CustomDocumentPropertiesDropDown);
            applicationEventHandler.SubscribeEventHandler(eventHandler);
        }

        // TODO Move to other class.
        private void OnDocumentChange()
        {
            // It is safe to access the Microsoft Word Application owner if the
            // DocumentChange event occurs. We don't always need to retrieve the
            // owner if we need it, since the Microsoft Word Application owner
            // exists as long as this Add-in runs.
            this.ApplicationWindow = ProcessUtils.MainWindowWin32HandleOfCurrentProcess();
        }

        // TODO This is completely static, therefore we do need a custom provider for each XML dtd.
        // Other possibility: Simple Key value XML?
        // OR: Let the user specifiy the XPath, e.g. /*/subsystems/subsystem[1]/components/component[1]/parameters would return all parameters for the first component of the first subsystem. But then we do need to specify the XPath for each column of the table, eg. child:://propertyName for parameters.
        private void OnClick_ButtonBindCustomXMLPart(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = this.application.ActiveDocument;
            CustomXMLPartRepository repository = new CustomXMLPartRepository(
                document.CustomXMLParts);

            this.CustomXMLPartsDialog.PopulateCustomXMLPartsListView(repository.FindNotBuiltIn());
            this.CustomXMLPartsDialog.ShowDialog(this.ApplicationWindow);

            if (DialogResult.Cancel.Equals(this.CustomXMLPartsDialog.DialogResult))
            {
                return;
            }

            string customXMLPartID = this.CustomXMLPartsDialog.LastSelectedCustomXMLPartID;
            Office.CustomXMLPart customXMLPart = repository.FindByID(customXMLPartID);

            Word.ContentControl contentControl = this.application.Selection.Range.ContentControls[1];
            Word.Range range = contentControl.Range;
            Word.Table table = range.Tables[1];

            Office.CustomXMLNodes subsystems = customXMLPart.SelectNodes("/ns0:defaultsystemparameters/ns0:subsystems/ns0:subsystem");

            // The bridge to OOXML.
            Microsoft.Office.Tools.Word.Document extendedDocument = Globals.Factory.GetVstoObject(document);

            this.application.ScreenUpdating = false;
            ProgressForm progressForm = new ProgressForm();
            progressForm.ChangeLabelText("Processing current document. Please wait and do not close Microsoft Word...");
            progressForm.Show(this.ApplicationWindow);

            foreach (Office.CustomXMLNode subsystemNode in subsystems)
            {
                string systemName = subsystemNode.SelectSingleNode("child::ns0:propertyName").Text;

                Office.CustomXMLNodes components = subsystemNode.SelectNodes("child::ns0:components/ns0:component");

                foreach (Office.CustomXMLNode componentNode in components)
                {
                    string componentName = componentNode.SelectSingleNode("child::ns0:propertyName").Text;

                    Office.CustomXMLNodes parameters = componentNode.SelectNodes("child::ns0:parameters/ns0:parameter");

                    foreach (Office.CustomXMLNode parameterNode in parameters)
                    {
                        table.Rows.Add();

                        Office.CustomXMLNode attributeNode = parameterNode.Attributes[1];

                        table.Cell(table.Rows.Count, table.Columns.Count - 2).Range.Select();
                        Microsoft.Office.Tools.Word.ContentControl checkBoxControl = extendedDocument.Controls.AddContentControl(attributeNode.XPath, Word.WdContentControlType.wdContentControlCheckBox);
                        checkBoxControl.Checked = Convert.ToBoolean(attributeNode.Text);
                        checkBoxControl.LockContentControl = true;
                        checkBoxControl.LockContents = true;

                        Office.CustomXMLNode keyNode = parameterNode.SelectSingleNode("child::ns0:propertyName");
                        table.Cell(table.Rows.Count, table.Columns.Count - 1).Range.Select();
                        
                        // TODO Name?! WTF how to automate that?!
                        Microsoft.Office.Tools.Word.PlainTextContentControl plainTextControl = extendedDocument.Controls.AddPlainTextContentControl(keyNode.XPath);
                        plainTextControl.XMLMapping.SetMappingByNode(keyNode);
                        plainTextControl.LockContentControl = true;

                        // TODO Causes SystemAccessViolation. Why?
                        // plainTextControl.LockContents = true;
                        Office.CustomXMLNode valueNode = parameterNode.SelectSingleNode("child::ns0:value");
                        if (null != valueNode)
                        {
                            table.Cell(table.Rows.Count, table.Columns.Count).Range.Select();
                            Microsoft.Office.Tools.Word.PlainTextContentControl plainTextControlValue = extendedDocument.Controls.AddPlainTextContentControl(valueNode.XPath);
                            plainTextControlValue.XMLMapping.SetMappingByNode(valueNode);
                            plainTextControlValue.LockContentControl = true;
                            plainTextControlValue.LockContents = true;
                        }
                    }
                }
            }

            progressForm.Close();
            this.application.ScreenUpdating = true;
        }

        private void OnClick_ToggleButtonShowFieldCodes(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;
            this.application.ActiveWindow.View.ShowFieldCodes = toggleButton.Checked;
        }

        // TODO The button state isn't in sync if the option is set via another method.
        private void OnClick_ToggleButtonShowFieldShading(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;
            this.application.ActiveDocument.FormFields.Shaded = toggleButton.Checked;
        }

        // TODO The button state isn't in sync if the option is set via another method.
        private void OnSelectionChanged_DropDownFieldShading(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown dropDown = (RibbonDropDown)sender;
            RibbonDropDownItem dropDownItem = dropDown.SelectedItem;
            Type enumType = typeof(Word.WdFieldShading);
            string enumValue = dropDownItem.Tag.ToString();

            Word.WdFieldShading fieldShading = (Word.WdFieldShading)Enum.Parse(enumType, enumValue);
            this.application.ActiveWindow.View.FieldShading = fieldShading;
        }

        private void OnClick_CheckBoxHideInternal(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox checkBox = (RibbonCheckBox)sender;
            this.CustomDocumentPropertiesDropDown.Update(checkBox.Checked);
        }

        private void OnClick_ButtonIncludeText(object sender, RibbonControlEventArgs e)
        {
            new InsertFileDialog(
                this.application,
                this.FieldFactory,
                Settings.Default.DocPropertyNameForLastDirectoryPath).Show();
        }

        private void OnClick_ButtonIncludePicture(object sender, RibbonControlEventArgs e)
        {
            // TODO It does not seem to be possible to specify a default path for a built-in dialog.
            // A possible solution would be, to replace the built-in dialog with a custom Windows form.
            // http://answers.microsoft.com//office/forum/office_2007-word/ms-word-defaults-to-a-set-folder-under-the/d604a81e-aa68-44e9-b7e0-ca9ad8f17e33
            new InsertPictureDialog(
                this.application,
                this.FieldFactory,
                Settings.Default.DocPropertyNameForLastDirectoryPath).Show();
        }

        private void OnClick_ButtonIncludeObject(object sender, RibbonControlEventArgs e)
        {
            new InsertObjectDialog(this.application).Show();
        }

        private void OnClick_ButtonCompareDocuments(object sender, RibbonControlEventArgs e)
        {
            new CompareDocumentsDialog(this.application).Show();
        }

        #region Event handler to show forms.

        private void OnClick_ButtonShowAboutForm(object sender, RibbonControlEventArgs e)
        {
            this.AboutDialog.ShowDialog(this.ApplicationWindow);
        }

        private void OnClick_ButtonShowConfigurationForm(object sender, RibbonControlEventArgs e)
        {
            this.ConfigurationDialog.ShowDialog(this.ApplicationWindow);
        }

        private void OnClick_ButtonShowHelpForm(object sender, RibbonControlEventArgs e)
        {
            this.HelpDialog.ShowDialog(this.ApplicationWindow);
        }

        private void OnClick_ToggleButtonShowMessagesForm(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;

            if (toggleButton.Checked)
            {
                this.messagesForm.Show(this.ApplicationWindow);
            }
            else
            {
                this.messagesForm.Visible = toggleButton.Checked;
            }
        }

        #endregion

        private void OnSelectionChanged_DropDownCustomDocumentProperties(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown dropDown = (RibbonDropDown)sender;

            string propertyName = dropDown.SelectedItem.Label;
            bool mergeFormat = this.toggleButtonFieldFormatMergeFormat.Checked;

            this.FieldFactory.InsertDocProperty(propertyName, mergeFormat);

            dropDown.SelectedItemIndex = 0;
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert
        /// Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_SplitButtonInsertField(object sender, RibbonControlEventArgs e)
        {
            new InsertFieldDialog(this.application).Show();
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert Empty
        /// Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ButtonInsertEmptyField(object sender, RibbonControlEventArgs e)
        {
            this.FieldFactory.InsertEmpty(
                this.application.Selection.Range.Text,
                this.toggleButtonFieldFormatMergeFormat.Checked);
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert Date
        /// Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ButtonInsertDateField(object sender, RibbonControlEventArgs e)
        {
            this.FieldFactory.InsertDate(this.toggleButtonFieldFormatMergeFormat.Checked);
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert Time
        /// Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ButtonInsertTimeField(object sender, RibbonControlEventArgs e)
        {
            this.FieldFactory.InsertTime(this.toggleButtonFieldFormatMergeFormat.Checked);
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert List
        /// Number Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ButtonInsertListNumField(object sender, RibbonControlEventArgs e)
        {
            this.FieldFactory.InsertListNum(this.toggleButtonFieldFormatMergeFormat.Checked);
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert Page
        /// Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ButtonInsertPageField(object sender, RibbonControlEventArgs e)
        {
            this.FieldFactory.InsertPage(this.toggleButtonFieldFormatMergeFormat.Checked);
        }

        /// <summary>
        /// Handles the <i>Click</i> event for all toggle buttons related to
        /// field code formatting.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ToggleButtonFieldFormat(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton checkBox = (RibbonToggleButton)sender;

            // ATTENTION: This only works by following the following conventions:
            // * The sender must be a RibbonToggleButton.
            // * The Id of the RibbonToggleButton must start with "toggleButtonFieldFormat".
            // * The Id must end with the field format name, e.g. "MergeFormat" or "OrdText".
            string switchName = e.Control.Id.Replace("toggleButtonFieldFormat", string.Empty);

            FieldFormatSwitch fieldSwitch = new FieldFormatSwitch(switchName);
            Word.Field selectedField = this.application.Selection.SelectedFields().ElementAt(0);

            if (checkBox.Checked)
            {
                fieldSwitch.AddToField(selectedField);
            }
            else
            {
                fieldSwitch.RemoveFromField(selectedField);
            }

            // Updates the field and preserves the UI state of the field.
            bool fieldCodeVisible = selectedField.ShowCodes;
            this.application.ScreenUpdating = !fieldCodeVisible;
            selectedField.Update();
            selectedField.ShowCodes = fieldCodeVisible;
            this.application.ScreenUpdating = fieldCodeVisible;
        }

        private void OnClick_ButtonUpdateFromSource(object sender, RibbonControlEventArgs e)
        {
            IEnumerable<Word.Field> fields = this.application.Selection.SelectedIncludeTextFields();
            int fieldCount = fields.Count();

            if (DialogResult.Yes == MessageBoxes.ShowMessageBoxWhetherToUpdateContentFromSource(fieldCount))
            {
                new FieldUpdater(fields, new UpdateTarget()).Update();
            }
        }

        private void OnClick_ButtonUpdateToSource(object sender, RibbonControlEventArgs e)
        {
            IEnumerable<Word.Field> fields = this.application.Selection.SelectedIncludeTextFields();
            int fieldCount = fields.Count();
            string filePath = new IncludeField(fields.ElementAt(0)).FilePath;

            if (new FileInfo(filePath).IsReadOnly)
            {
                MessageBoxes.ShowMessageBoxFileIsReadOnly(filePath);
            }
            else if (DialogResult.Yes == MessageBoxes.ShowMessageBoxWhetherToUpdateContentInSource(fieldCount))
            {
                new FieldUpdater(fields, new UpdateSource()).Update();
            }
        }

        private void OnClick_ButtonCreateCustomDocumentProperty(object sender, RibbonControlEventArgs e)
        {
            CustomDocumentPropertyReader customDocumentPropertyReader = new CustomDocumentPropertyReader(this.application.ActiveDocument);
            CustomDocumentPropertyWriter customDocumentPropertyWriter = new CustomDocumentPropertyWriter(this.application.ActiveDocument);

            // TODO Referencing Visual Basic (VB) is ugly, but does the job.
            // Further developments should implement a custom form which also
            // allows to specifiy the data type of the custom document property.
            string propertyName = VB.Interaction.InputBox(
                "Enter the name of the custom property to set.",
                "Write a custom property");

            if (string.Empty == propertyName)
            {
                MessageBoxes.ShowMessageBoxNoCustomDocumentPropertyModfied();
                return;
            }

            if (customDocumentPropertyReader.Exists(propertyName))
            {
                if (DialogResult.No == MessageBoxes.ShowMessageBoxWhetherToOverwriteCustomDocumentProperty(propertyName))
                {
                    MessageBoxes.ShowMessageBoxNoCustomDocumentPropertyModfied();
                    return;
                }
            }

            string propertyValue = VB.Interaction.InputBox(
                "Enter the value of the property with the name '" + propertyName + "' .",
                "Write a custom property");

            if (string.Empty == propertyValue)
            {
                MessageBoxes.ShowMessageBoxNoCustomDocumentPropertyModfied();
                return;
            }

            customDocumentPropertyWriter.Set(propertyName, propertyValue);

            this.CustomDocumentPropertiesDropDown.Update(this.checkBoxHideInternal.Checked);

            MessageBoxes.ShowMessageBoxSetCustomDocumentPropertySuccess(propertyName, propertyValue);
        }

        private void OnClick_ButtonOpenSourceFile(object sender, RibbonControlEventArgs e)
        {
            // Open each referenced file (e.g. a Microsoft Word document) in the current selection.
            foreach (Word.Field field in this.application.Selection.SelectedIncludeTextFields())
            {
                Process.Start(new IncludeField(field).FilePath);
            }
        }

        private void OnClick_ButtonCheckReferences(object sender, RibbonControlEventArgs e)
        {
            string filePath;
            string lastModifiedActual;
            string lastModifiedExpected;

            // http://pmueller.de/blog/word2007grafik.html
            // Word Bug: http://stackoverflow.com/questions/17109200/ms-word-includepicture-field-code
            IList<Word.Field> fields = this.application.Selection.IncludeFields().ToList();

            foreach (Word.Field field in fields)
            {
                filePath = new IncludeField(field).FilePath;
                lastModifiedExpected = File.GetLastWriteTimeUtc(filePath).ToString("u");

                if (null == field.Next || null == field.Next.Next || !field.Next.Next.Type.Equals(Word.WdFieldType.wdFieldEmpty))
                {
                    MessageBox.Show(
                        "An error occured while parsing the field code. Ensure that the field has been created via " + Settings.Default.ApplicationName + ".",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    continue;
                }

                lastModifiedActual = field.Next.Next.Code.Text.Trim();

                if (lastModifiedExpected != lastModifiedActual)
                {
                    MessageBox.Show(
                        "The referenced source file " + filePath + " has been modified since it has been included in this target document.",
                        "Question",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    // TODO How can we solve a possible merge conflict?
                }
            }

            MessageBox.Show(
                "No problems have been found by " + Settings.Default.ApplicationName + " in the current selection of this document.",
                "Information",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void OnClick_ButtonFieldUpdate(object sender, RibbonControlEventArgs e)
        {
            // Update each field in the current selection.
            new FieldUpdater(
                this.application.Selection.SelectedFields(),
                new UpdateTarget()).Update();
        }

        private void OnClick_ToggleButtonFieldLock(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;
            bool lockField = false;

            foreach (Word.Field field in this.application.Selection.SelectedFields())
            {
                lockField = toggleButton.Checked;
                field.Locked = lockField;
                this.buttonFieldUpdate.Enabled = !lockField;
            }
        }

        private void OnClick_ToggleButtonShowFieldCode(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;

            foreach (Word.Field field in this.application.Selection.SelectedFields())
            {
                field.ShowCodes = toggleButton.Checked;
            }
        }
    }
}
