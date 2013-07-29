//------------------------------------------------------------------------------
// <copyright file="ConfigurationForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;

    internal partial class ConfigurationForm : Form
    {
        private readonly Settings settings;

        public ConfigurationForm(Settings settings)
        {
            this.settings = settings;
            this.InitializeComponent();
            this.FormClosing += this.OnFormClosing;
            this.Load += this.OnLoad;
        }

        private void OnLoad(object sender, EventArgs e)
        {
            this.textBoxGraphicsDirectoryName.Text = this.settings.GraphicsDirectoryName;
            this.textBoxXMLDirectoryName.Text = this.settings.XMLDirectoryName;
            this.textBoxWordTemplateFilePrefix.Text = this.settings.WordTemplateFilename;
            this.textBoxWordTemplateFileExtensions.Text = this.settings.WordTemplateFileExtensions;
        }

        private void OnFormClosing(object sender, EventArgs e)
        {
            this.settings.WordTemplateFilename = this.textBoxWordTemplateFilePrefix.Text.ToString();
            this.settings.WordTemplateFileExtensions = this.textBoxWordTemplateFileExtensions.Text.ToString();
            this.settings.XMLDirectoryName = this.textBoxXMLDirectoryName.Text.ToString();

            // TODO Add validation of configuration settings.
            this.settings.Save();
        }
    }
}
