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
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Commands;
    using FlorianWolters.Office.Word.AddIn.CBA.CustomXML;
    using FlorianWolters.Office.Word.AddIn.CBA.EventHandlers;
    using FlorianWolters.Office.Word.AddIn.CBA.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Office.Word.AddIn.ComponentAddIn.Commands;
    using FlorianWolters.Office.Word.Commands;
    using FlorianWolters.Office.Word.Dialogs;
    using FlorianWolters.Office.Word.DocumentProperties;
    using FlorianWolters.Office.Word.Event;
    using FlorianWolters.Office.Word.Event.EventExceptionHandlers;
    using FlorianWolters.Office.Word.Event.EventHandlers;
    using FlorianWolters.Office.Word.Extensions;
    using FlorianWolters.Office.Word.Fields;
    using FlorianWolters.Office.Word.Fields.Switches;
    using FlorianWolters.Office.Word.Fields.UpdateStrategies;
    using FlorianWolters.Reflection;
    using FlorianWolters.Windows.Forms;
    using Microsoft.Office.Tools.Ribbon;
    using NLog;
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
        /// The Logger for the class <see cref="Ribbon"/>.
        /// </summary>
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// Gets or sets the <see cref="Word.Application"/>, this <see
        /// cref="Ribbon"/> is running in.
        /// </summary>
        private Word.Application Application { get; set; }

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
        /// Occurs when <see cref="Ribbon"/> is loaded into the Microsoft
        /// Office Application. 
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            Logger.Debug(System.Reflection.MethodBase.GetCurrentMethod().Name);

            // TODO This is the suggested way to access the Word application?!
            // TODO Could be moved to the constructor, but the constructor is
            // defined in Ribbon.Designer.cs, which is modified by the Designer.
            this.Application = Globals.ThisAddIn.Application;

            CustomDocumentPropertyReader customDocumentPropertyReader = new CustomDocumentPropertyReader();

            this.FieldFactory = new FieldFactory(this.Application, customDocumentPropertyReader);
            this.CustomDocumentPropertiesDropDown = new CustomDocumentPropertiesDropDown(
                this.Application,
                this.Factory,
                this.dropDownCustomDocumentProperties,
                customDocumentPropertyReader);
            this.InitializeForms();
            this.RegisterEventHandler();
        }

        /// <summary>
        /// Initializes the forms used by this <see cref="Ribbon"/>.
        /// </summary>
        private void InitializeForms()
        {
            AssemblyInfo assemblyInfo = new AssemblyInfo(Assembly.GetExecutingAssembly());

            // TODO Validate configuration options.
            Settings defaultSettings = Settings.Default;

            this.ConfigurationDialog = new ConfigurationForm(defaultSettings);
            this.AboutDialog = new AboutForm(assemblyInfo, defaultSettings);
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
                this.buttonShowHelp.Enabled = false;
            }
        }

        /// <summary>
        /// Occurs when <see cref="Ribbon"/> instance is closing.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">An object that contains no event data.</param>
        private void Ribbon_Close(object sender, EventArgs e)
        {
            Logger.Debug(System.Reflection.MethodBase.GetCurrentMethod().Name);
        }

        // TODO How to configure this?

        /// <summary>
        /// Registers all <i>Event Handlers</i> for this <see cref="Ribbon"/>.
        /// </summary>
        private void RegisterEventHandler()
        {
            // I am not sure if the implemented approach is the "best" one, since it contains a lot of boilerplate code.
            // - "Event Handler" objects are created via the Factory Method creational design searchPattern.
            // - The Factory Method also creates the correct "Command" object and injects it into the "Event Handler".
            // - The Factory Method registers the "Event Handler" at the "Global Event Handler".

            // A "Command" does not know anything about a "Event Handler". A "Command" implements business logic related to a Word.Application ONLY.
            // A "Event Handler" class can implement Interfaces which do contain the signature for the events in the Word object model.
            // Every implementation of that method calls the "HandleResult" method of the "Command" object.

            // ATTENTION: Since we always inject the Word.Application into the objects, we can always access the current state of the Microsoft Word application.
            // If we would work with Word.Document, we would always have to make sure, that the reference is up-to-date.
            ApplicationEventHandler applicationEventHandler = new ApplicationEventHandler(this.Application);
            IEventExceptionHandler eventExceptionHandler = new MessageBoxEventExceptionHandler();

            // TODO Allow configuration of event handlers and simplify creation.
            this.InitializeRefreshCustomXMLParts(applicationEventHandler, eventExceptionHandler);
            this.InitializeWriteCustomDocumentProperties(applicationEventHandler);
            applicationEventHandler.SubscribeEventHandler(
                new UpdateAttachedTemplateEventHandler(
                    new UpdateAttachedTemplateCommand(this.Application),
                    eventExceptionHandler));

            // The RibbonStateEventHandler ensures that the state of the UI of this Ribbon is correctly set.
            IEventHandler eventHandler = new RibbonStateEventHandler(
                this.Application,
                this,
                this.CustomDocumentPropertiesDropDown);
            applicationEventHandler.SubscribeEventHandler(eventHandler);
        }

        private void InitializeRefreshCustomXMLParts(
            ApplicationEventHandler applicationEventHandler,
            IEventExceptionHandler eventExceptionHandler)
        {
            ICommand command = new RefreshCustomXMLPartsCommand(applicationEventHandler.Application);
            IEventHandler eventHandler = new RefreshCustomXMLPartsCommandEventHandler(command, eventExceptionHandler);
            applicationEventHandler.SubscribeEventHandler(eventHandler);
        }

        private void InitializeWriteCustomDocumentProperties(ApplicationEventHandler applicationEventHandler)
        {
            IEventHandler eventHandler = new WriteCustomDocumentPropertiesEventHandler(applicationEventHandler.Application);
            applicationEventHandler.SubscribeEventHandler(eventHandler);
        }

        // TODO Remove or move to other classes.
        private void OnDocumentChange()
        {
            Logger.Debug(System.Reflection.MethodBase.GetCurrentMethod().Name);

            // It is safe to access the Microsoft Word Application window if the
            // DocumentChange event occurs. We don't always need to retrieve the
            // window if we need it, since the Microsoft Word Application window
            // exists as long as this Add-in runs.
            this.ApplicationWindow = ProcessUtils.MainWindowWin32HandleOfCurrentProcess();

            if (this.Application.HasOpenDocuments())
            {
                Word.Document activeDocument = this.Application.ActiveDocument;

                if (activeDocument.IsSaved())
                {
                    // All operations that require that the active document is saved.
                }
            }
        }

        // TODO This is completely static, therefore we do need a custom provider for each XML dtd.
        // Other possibility: Simple Key value XML?
        // OR: Let the user specifiy the XPath, e.g. /*/subsystems/subsystem[1]/components/component[1]/parameters would return all parameters for the first component of the first subsystem. But then we do need to specify the XPath for each column of the table, eg. child:://propertyName for parameters.
        private void OnClick_ButtonBindCustomXMLPart(object sender, RibbonControlEventArgs e)
        {
            Word.Document document = this.Application.ActiveDocument;
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

            Word.ContentControl contentControl = this.Application.Selection.Range.ContentControls[1];
            Word.Range range = contentControl.Range;
            Word.Table table = range.Tables[1];

            Office.CustomXMLNodes subsystems = customXMLPart.SelectNodes("/ns0:defaultsystemparameters/ns0:subsystems/ns0:subsystem");

            // The bridge to OOXML.
            Microsoft.Office.Tools.Word.Document extendedDocument = Globals.Factory.GetVstoObject(document);

            this.Application.ScreenUpdating = false;
            ProgressForm progressForm = new ProgressForm();
            progressForm.ChangeLabelText("Processing current Document. Please wait and do not close Microsoft Word...");
            progressForm.Show(this.ApplicationWindow);

            foreach (Office.CustomXMLNode subsystemNode in subsystems)
            {
                string systemName = subsystemNode.SelectSingleNode("child::ns0:propertyName").Text;
                Logger.Trace("System: " + systemName);

                Office.CustomXMLNodes components = subsystemNode.SelectNodes("child::ns0:components/ns0:component");

                foreach (Office.CustomXMLNode componentNode in components)
                {
                    string componentName = componentNode.SelectSingleNode("child::ns0:propertyName").Text;
                    Logger.Trace("Component: " + componentName);

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
            this.Application.ScreenUpdating = true;
        }

        private void OnClick_ToggleButtonShowFieldCodes(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;
            this.Application.ActiveWindow.View.ShowFieldCodes = toggleButton.Checked;
        }

        // TODO The button state isn't in sync if the option is set via another method.
        private void OnClick_ToggleButtonShowFieldShading(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton toggleButton = (RibbonToggleButton)sender;
            this.Application.ActiveDocument.FormFields.Shaded = toggleButton.Checked;
        }

        // TODO The button state isn't in sync if the option is set via another method.
        private void OnSelectionChanged_DropDownFieldShading(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown dropDown = (RibbonDropDown)sender;
            RibbonDropDownItem dropDownItem = dropDown.SelectedItem;
            Type enumType = typeof(Word.WdFieldShading);
            string enumValue = dropDownItem.Tag.ToString();

            Word.WdFieldShading fieldShading = (Word.WdFieldShading)Enum.Parse(enumType, enumValue);
            this.Application.ActiveWindow.View.FieldShading = fieldShading;
        }

        private void OnClick_CheckBoxHideInternal(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox checkBox = (RibbonCheckBox)sender;
            this.CustomDocumentPropertiesDropDown.Update(checkBox.Checked);
        }

        private void OnClick_ButtonIncludeText(object sender, RibbonControlEventArgs e)
        {
            new InsertFileDialog(
                this.Application,
                this.FieldFactory,
                Settings.Default.DocPropertyNameForLastDirectoryPath).Show();
        }

        private void OnClick_ButtonIncludePicture(object sender, RibbonControlEventArgs e)
        {
            new InsertPictureDialog(
                this.Application,
                this.FieldFactory,
                Settings.Default.DocPropertyNameForLastDirectoryPath).Show();
        }

        private void OnClick_ButtonIncludeObject(object sender, RibbonControlEventArgs e)
        {
            new InsertObjectDialog(this.Application).Show();
        }

        private void OnClick_ButtonCompareDocuments(object sender, RibbonControlEventArgs e)
        {
            new CompareDocumentsDialog(this.Application).Show();
        }

        private void OnClick_ButtonShowAboutDialog(object sender, RibbonControlEventArgs e)
        {
            this.AboutDialog.ShowDialog(this.ApplicationWindow);
        }

        private void OnClick_ButtonShowConfigurationDialog(object sender, RibbonControlEventArgs e)
        {
            this.ConfigurationDialog.ShowDialog(this.ApplicationWindow);
        }

        private void OnClick_ButtonShowHelp(object sender, RibbonControlEventArgs e)
        {
            this.HelpDialog.ShowDialog();
        }

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
            new InsertFieldDialog(this.Application).Show();
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the split button <i>Insert Empty
        /// Field</i>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_ButtonInsertEmptyField(object sender, RibbonControlEventArgs e)
        {
            this.FieldFactory.InsertEmpty(this.toggleButtonFieldFormatMergeFormat.Checked);
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
            Word.Field selectedField = this.Application.Selection.SelectedFields().ElementAt(0);

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
            this.Application.ScreenUpdating = !fieldCodeVisible;
            selectedField.Update();
            selectedField.ShowCodes = fieldCodeVisible;
            this.Application.ScreenUpdating = fieldCodeVisible;
        }

        private void OnClick_ButtonUpdateFromSource(object sender, RibbonControlEventArgs e)
        {
            IEnumerable<Word.Field> fields = this.Application.Selection.SelectedIncludeFields();
            int fieldCount = fields.Count();

            if (DialogResult.Yes == MessageBoxes.ShowMessageBoxWhetherToUpdateContentFromSource(fieldCount))
            {
                new FieldUpdater(fields, new UpdateTarget()).Update();
            }
        }

        private void OnClick_ButtonUpdateToSource(object sender, RibbonControlEventArgs e)
        {
            IEnumerable<Word.Field> fields = this.Application.Selection.SelectedIncludeTextFields();
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
            CustomDocumentPropertyReader customDocumentPropertyReader = new CustomDocumentPropertyReader(this.Application.ActiveDocument);
            CustomDocumentPropertyWriter customDocumentPropertyWriter = new CustomDocumentPropertyWriter(this.Application.ActiveDocument);

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
            IEnumerable<Word.Field> fields = this.Application.Selection.SelectedIncludeFields();

            foreach (Word.Field field in fields)
            {
                Process.Start(new IncludeField(field).FilePath);
            }
        }

        private void OnClick_ButtonCheckReferences(object sender, RibbonControlEventArgs e)
        {
            // TODO
        }
    }
}
