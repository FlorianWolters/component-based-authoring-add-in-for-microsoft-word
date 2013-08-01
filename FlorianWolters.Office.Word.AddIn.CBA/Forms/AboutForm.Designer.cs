﻿//------------------------------------------------------------------------------
// <copyright file="AboutForm.Designer.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
// <auto-generated/>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    internal partial class AboutForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.TextBox textBoxVersion;
        private System.Windows.Forms.TextBox textBoxAuthor;
        private System.Windows.Forms.TextBox textBoxDescription;
        private System.Windows.Forms.PictureBox pictureBoxHostingService;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.Label labelVersion;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutForm));
            System.Windows.Forms.Label labelAuthor;
            System.Windows.Forms.Label labelName;
            System.Windows.Forms.Label labelContribute;
            System.Windows.Forms.Label labelDescription;
            this.textBoxAuthor = new System.Windows.Forms.TextBox();
            this.textBoxVersion = new System.Windows.Forms.TextBox();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.pictureBoxHostingService = new System.Windows.Forms.PictureBox();
            this.textBoxDescription = new System.Windows.Forms.TextBox();
            labelVersion = new System.Windows.Forms.Label();
            labelAuthor = new System.Windows.Forms.Label();
            labelName = new System.Windows.Forms.Label();
            labelContribute = new System.Windows.Forms.Label();
            labelDescription = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHostingService)).BeginInit();
            this.SuspendLayout();
            // 
            // labelVersion
            // 
            resources.ApplyResources(labelVersion, "labelVersion");
            labelVersion.Name = "labelVersion";
            // 
            // labelAuthor
            // 
            resources.ApplyResources(labelAuthor, "labelAuthor");
            labelAuthor.Name = "labelAuthor";
            // 
            // labelName
            // 
            resources.ApplyResources(labelName, "labelName");
            labelName.Name = "labelName";
            // 
            // labelContribute
            // 
            resources.ApplyResources(labelContribute, "labelContribute");
            labelContribute.Name = "labelContribute";
            // 
            // labelDescription
            // 
            resources.ApplyResources(labelDescription, "labelDescription");
            labelDescription.Name = "labelDescription";
            // 
            // textBoxAuthor
            // 
            resources.ApplyResources(this.textBoxAuthor, "textBoxAuthor");
            this.textBoxAuthor.Name = "textBoxAuthor";
            this.textBoxAuthor.ReadOnly = true;
            // 
            // textBoxVersion
            // 
            resources.ApplyResources(this.textBoxVersion, "textBoxVersion");
            this.textBoxVersion.Name = "textBoxVersion";
            this.textBoxVersion.ReadOnly = true;
            // 
            // textBoxName
            // 
            resources.ApplyResources(this.textBoxName, "textBoxName");
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.ReadOnly = true;
            // 
            // pictureBoxHostingService
            // 
            this.pictureBoxHostingService.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxHostingService.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBoxHostingService.Image = global::FlorianWolters.Office.Word.AddIn.CBA.Properties.Resources.GitHub_Logo;
            resources.ApplyResources(this.pictureBoxHostingService, "pictureBoxHostingService");
            this.pictureBoxHostingService.Name = "pictureBoxHostingService";
            this.pictureBoxHostingService.TabStop = false;
            this.pictureBoxHostingService.Click += new System.EventHandler(this.OnClick_PictureBoxHostingService);
            this.pictureBoxHostingService.MouseHover += new System.EventHandler(this.OnMouseHover_PictureBoxHostingService);
            // 
            // textBoxDescription
            // 
            resources.ApplyResources(this.textBoxDescription, "textBoxDescription");
            this.textBoxDescription.Name = "textBoxDescription";
            this.textBoxDescription.ReadOnly = true;
            // 
            // AboutForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(labelDescription);
            this.Controls.Add(this.textBoxDescription);
            this.Controls.Add(labelContribute);
            this.Controls.Add(this.pictureBoxHostingService);
            this.Controls.Add(this.textBoxName);
            this.Controls.Add(labelName);
            this.Controls.Add(this.textBoxVersion);
            this.Controls.Add(this.textBoxAuthor);
            this.Controls.Add(labelAuthor);
            this.Controls.Add(labelVersion);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutForm";
            this.ShowInTaskbar = false;
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxHostingService)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}
