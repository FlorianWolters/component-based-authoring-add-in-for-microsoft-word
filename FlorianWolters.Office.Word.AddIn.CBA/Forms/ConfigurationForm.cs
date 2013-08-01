//------------------------------------------------------------------------------
// <copyright file="ConfigurationForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;

    /// <summary>
    /// The class <see cref="ConfigurationForm"/> allows to configure the <see
    /// cref="Settings"/> of the Microsoft Word Application-Level Add-In.
    /// </summary>
    internal partial class ConfigurationForm : Form
    {
        /// <summary>
        /// The <see cref="Settings"/> of the Microsoft Word Application-Level Add-In.
        /// </summary>
        private readonly Settings settings;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConfigurationForm"/> class with the specified <see cref="Settings"/> to display.
        /// </summary>
        /// <param name="settings">The <see cref="Settings"/> of the Microsoft Word Application-Level Add-In.
        /// </param>
        public ConfigurationForm(Settings settings)
        {
            this.InitializeComponent();
            this.settings = settings;
        }

        /// <summary>
        /// Occurs before this <see cref="ConfigurationForm"/> is displayed for the first time.
        /// <para>
        /// Writes data of the <see cref="Settings"/> to the form controls.
        /// </para>
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        private void OnLoad(object sender, EventArgs e)
        {
            this.textBoxGraphicsDirectoryName.Text = this.settings.GraphicsDirectoryName;
            this.textBoxXMLDirectoryName.Text = this.settings.XMLDirectoryName;
            this.textBoxWordTemplateFilePrefix.Text = this.settings.WordTemplateFilename;
            this.textBoxWordTemplateFileExtensions.Text = this.settings.WordTemplateFileExtensions;

            this.checkBoxUpdateStylesOnOpen.Checked = this.settings.ActivateUpdateStylesOnOpen;
            this.checkBoxRefreshCustomXMLParts.Checked = this.settings.RefreshCustomXMLParts;
            this.checkBoxUpdateAttachedTemplate.Checked = this.settings.UpdateAttachedTemplate;
            this.checkBoxUpdateFields.Checked = this.settings.UpdateFields;
            this.checkBoxWriteCustomDocumentProperties.Checked = this.settings.WriteCustomDocumentProperties;
        }

        /// <summary>
        /// Occurs before this <see cref="ConfigurationForm"/> is closed.
        /// <para>
        /// Writes data from the form controls to the <see cref="Settings"/>.
        /// </para>
        /// </summary>
        /// <param name="sender">The sender of the event.</param>
        /// <param name="e">A <see cref="FormClosingEventArgs"/> that contains the event data.</param>
        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult.Equals(DialogResult.OK))
            {
                // TODO Refactpr validation of configuration settings (move to other class).

                string graphicsDirectoryName = this.textBoxGraphicsDirectoryName.Text.ToString();
                string xmlDirectoryName = this.textBoxXMLDirectoryName.Text.ToString();
                string templateName = this.textBoxWordTemplateFilePrefix.Text.ToString();
                
                // Taken from http://regexlib.com/REDetails.aspx?regexp_id=85
                const string pattern = "^(?!^(PRN|AUX|CLOCK\\$|NUL|CON|COM\\d|LPT\\d|\\..*)(\\..+)?$)[^\\x00-\\x1f\\\\?*:\\\";|/]+$";
                
                if (!Regex.IsMatch(graphicsDirectoryName, pattern, RegexOptions.IgnoreCase))
                {
                    MessageBox.Show("Invalid directory name for the graphic files.", "Notice");
                    e.Cancel = true;
                }

                if (!Regex.IsMatch(xmlDirectoryName, pattern))
                {
                    MessageBox.Show("Invalid directory name for the XML files.", "Notice");
                    e.Cancel = true;
                }

                if (!Regex.IsMatch(templateName, pattern))
                {
                    MessageBox.Show("Invalid directory name for the template file.", "Notice");
                    e.Cancel = true;
                }

                if (!e.Cancel)
                {
                    this.settings.GraphicsDirectoryName = graphicsDirectoryName;
                    this.settings.XMLDirectoryName = xmlDirectoryName;
                    this.settings.WordTemplateFilename = templateName;
                    this.settings.WordTemplateFileExtensions = this.textBoxWordTemplateFileExtensions.Text.ToString();

                    this.settings.ActivateUpdateStylesOnOpen = this.checkBoxUpdateStylesOnOpen.Checked;
                    this.settings.RefreshCustomXMLParts = this.checkBoxRefreshCustomXMLParts.Checked;
                    this.settings.UpdateAttachedTemplate = this.checkBoxUpdateAttachedTemplate.Checked;
                    this.settings.UpdateFields = this.checkBoxUpdateFields.Checked;
                    this.settings.WriteCustomDocumentProperties = this.checkBoxWriteCustomDocumentProperties.Checked;

                    this.settings.Save();
                }
            }
        }
    }
}
