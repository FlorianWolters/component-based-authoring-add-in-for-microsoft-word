﻿//------------------------------------------------------------------------------
// <copyright file="ConfigurationForm.Designer.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    internal partial class ConfigurationForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox textBoxXMLDirectoryName;
        private System.Windows.Forms.TextBox textBoxWordTemplateFilePrefix;
        private System.Windows.Forms.TextBox textBoxWordTemplateFileExtensions;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing"><c>true</c> if managed resources should be disposed; otherwise, <c>false</c>.</param>
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
            System.Windows.Forms.Label labelTemplateFileName;
            System.Windows.Forms.Label labelXMLDirectoryName;
            System.Windows.Forms.Label labelTemplateFileExtensions;
            System.Windows.Forms.Label labelGraphicsDirectoryName;
            this.textBoxXMLDirectoryName = new System.Windows.Forms.TextBox();
            this.textBoxWordTemplateFilePrefix = new System.Windows.Forms.TextBox();
            this.textBoxWordTemplateFileExtensions = new System.Windows.Forms.TextBox();
            this.textBoxGraphicsDirectoryName = new System.Windows.Forms.TextBox();
            labelTemplateFileName = new System.Windows.Forms.Label();
            labelXMLDirectoryName = new System.Windows.Forms.Label();
            labelTemplateFileExtensions = new System.Windows.Forms.Label();
            labelGraphicsDirectoryName = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelTemplateFileName
            // 
            labelTemplateFileName.AutoSize = true;
            labelTemplateFileName.Location = new System.Drawing.Point(67, 69);
            labelTemplateFileName.Name = "labelTemplateFileName";
            labelTemplateFileName.Size = new System.Drawing.Size(99, 13);
            labelTemplateFileName.TabIndex = 2;
            labelTemplateFileName.Text = "Template Filename:";
            // 
            // textBoxXMLDirectoryName
            // 
            this.textBoxXMLDirectoryName.Location = new System.Drawing.Point(172, 40);
            this.textBoxXMLDirectoryName.Name = "textBoxXMLDirectoryName";
            this.textBoxXMLDirectoryName.Size = new System.Drawing.Size(100, 20);
            this.textBoxXMLDirectoryName.TabIndex = 3;
            // 
            // labelXMLDirectoryName
            // 
            labelXMLDirectoryName.AutoSize = true;
            labelXMLDirectoryName.Location = new System.Drawing.Point(58, 43);
            labelXMLDirectoryName.Name = "labelXMLDirectoryName";
            labelXMLDirectoryName.Size = new System.Drawing.Size(108, 13);
            labelXMLDirectoryName.TabIndex = 4;
            labelXMLDirectoryName.Text = "XML Directory Name:";
            // 
            // textBoxWordTemplateFilePrefix
            // 
            this.textBoxWordTemplateFilePrefix.Location = new System.Drawing.Point(172, 66);
            this.textBoxWordTemplateFilePrefix.Name = "textBoxWordTemplateFilePrefix";
            this.textBoxWordTemplateFilePrefix.Size = new System.Drawing.Size(100, 20);
            this.textBoxWordTemplateFilePrefix.TabIndex = 5;
            // 
            // labelTemplateFileExtensions
            // 
            labelTemplateFileExtensions.AutoSize = true;
            labelTemplateFileExtensions.Location = new System.Drawing.Point(39, 95);
            labelTemplateFileExtensions.Name = "labelTemplateFileExtensions";
            labelTemplateFileExtensions.Size = new System.Drawing.Size(127, 13);
            labelTemplateFileExtensions.TabIndex = 6;
            labelTemplateFileExtensions.Text = "Template File Extensions:";
            // 
            // textBoxWordTemplateFileExtensions
            // 
            this.textBoxWordTemplateFileExtensions.Location = new System.Drawing.Point(172, 92);
            this.textBoxWordTemplateFileExtensions.Name = "textBoxWordTemplateFileExtensions";
            this.textBoxWordTemplateFileExtensions.Size = new System.Drawing.Size(100, 20);
            this.textBoxWordTemplateFileExtensions.TabIndex = 7;
            // 
            // labelGraphicsDirectoryName
            // 
            labelGraphicsDirectoryName.AutoSize = true;
            labelGraphicsDirectoryName.Location = new System.Drawing.Point(38, 15);
            labelGraphicsDirectoryName.Name = "labelGraphicsDirectoryName";
            labelGraphicsDirectoryName.Size = new System.Drawing.Size(128, 13);
            labelGraphicsDirectoryName.TabIndex = 8;
            labelGraphicsDirectoryName.Text = "Graphics Directory Name:";
            // 
            // textBoxGraphicsDirectoryName
            // 
            this.textBoxGraphicsDirectoryName.Location = new System.Drawing.Point(172, 12);
            this.textBoxGraphicsDirectoryName.Name = "textBoxGraphicsDirectoryName";
            this.textBoxGraphicsDirectoryName.Size = new System.Drawing.Size(100, 20);
            this.textBoxGraphicsDirectoryName.TabIndex = 9;
            // 
            // ConfigurationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 122);
            this.Controls.Add(this.textBoxGraphicsDirectoryName);
            this.Controls.Add(labelGraphicsDirectoryName);
            this.Controls.Add(this.textBoxWordTemplateFileExtensions);
            this.Controls.Add(labelTemplateFileExtensions);
            this.Controls.Add(this.textBoxWordTemplateFilePrefix);
            this.Controls.Add(labelXMLDirectoryName);
            this.Controls.Add(this.textBoxXMLDirectoryName);
            this.Controls.Add(labelTemplateFileName);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConfigurationForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Configuration";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxGraphicsDirectoryName;
    }
}
