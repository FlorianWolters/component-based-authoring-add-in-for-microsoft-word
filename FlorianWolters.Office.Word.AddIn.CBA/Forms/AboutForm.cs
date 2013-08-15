//------------------------------------------------------------------------------
// <copyright file="AboutForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System.Diagnostics;
    using System.Windows.Forms;
    using FlorianWolters.Office.Word.AddIn.CBA.Properties;
    using FlorianWolters.Reflection;

    /// <summary>
    /// The class <see cref="AboutForm"/> implements a simple <i>About</i> dialog, which displays information (e.g. the
    /// version) of the application.
    /// </summary>
    internal partial class AboutForm : Form
    {
        /// <summary>
        /// The <see cref="Settings"/> of this application.
        /// </summary>
        private readonly Settings settings;

        /// <summary>
        /// Initializes a new instance of the <see cref="AboutForm"/> class with
        /// the specified <see cref="AssemblyInfo"/> and the specified <see
        /// cref="Settings"/>.
        /// </summary>
        /// <param name="assemblyInfo">The <see cref="AssemblyInfo"/>.</param>
        /// <param name="settings">The <see cref="Settings"/>.</param>
        public AboutForm(AssemblyInfo assemblyInfo, Settings settings)
        {
            this.InitializeComponent();

            this.textBoxName.Text = assemblyInfo.Title;
            this.textBoxVersion.Text = assemblyInfo.Version.ToString();
            this.textBoxAuthor.Text = assemblyInfo.Company;
            this.textBoxDescription.Text = assemblyInfo.Description;
            this.settings = settings;
        }

        /// <summary>
        /// Handles the <i>Click</i> event for the <see cref="PictureBox"/> of
        /// this <see cref="AboutForm"/>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnClick_PictureBoxHostingService(object sender, System.EventArgs e)
        {
            Process.Start(this.settings.HostingServiceUrl);
        }

        /// <summary>
        /// Handles the <i>Mouse Hover</i> event for the <see
        /// cref="PictureBox"/> of this <see cref="AboutForm"/>.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The data for the event.</param>
        private void OnMouseHover_PictureBoxHostingService(object sender, System.EventArgs e)
        {
            PictureBox pictureBox = (PictureBox)sender;
            const string ToolTextLabel = "Click here to open the repository of the project on the web-based hosting service GitHub.";
            
            ToolTip toolTip = new ToolTip();
            toolTip.SetToolTip(pictureBox, ToolTextLabel);
        }
    }
}
