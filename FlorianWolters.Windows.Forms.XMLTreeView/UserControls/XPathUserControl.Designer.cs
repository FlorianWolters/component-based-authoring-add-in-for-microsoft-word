﻿//------------------------------------------------------------------------------
// <copyright file="XPathUserControl.Designer.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
// <auto-generated/>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.UserControls
{
    partial class XPathUserControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox textBoxXPathExpression;
        private System.Windows.Forms.TextBox textBoxXPathPrefixMapping;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        private void InitializeComponent()
        {
            System.Windows.Forms.GroupBox groupBoxXPath;
            System.Windows.Forms.Label labelXPathExpression;
            System.Windows.Forms.Label labelXPathPrefixMapping;
            this.textBoxXPathExpression = new System.Windows.Forms.TextBox();
            this.textBoxXPathPrefixMapping = new System.Windows.Forms.TextBox();
            groupBoxXPath = new System.Windows.Forms.GroupBox();
            labelXPathExpression = new System.Windows.Forms.Label();
            labelXPathPrefixMapping = new System.Windows.Forms.Label();
            groupBoxXPath.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxXPath
            // 
            groupBoxXPath.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            groupBoxXPath.Controls.Add(labelXPathExpression);
            groupBoxXPath.Controls.Add(labelXPathPrefixMapping);
            groupBoxXPath.Controls.Add(this.textBoxXPathExpression);
            groupBoxXPath.Controls.Add(this.textBoxXPathPrefixMapping);
            groupBoxXPath.Location = new System.Drawing.Point(0, 0);
            groupBoxXPath.Margin = new System.Windows.Forms.Padding(0);
            groupBoxXPath.Name = "groupBoxXPath";
            groupBoxXPath.Size = new System.Drawing.Size(400, 82);
            groupBoxXPath.TabIndex = 6;
            groupBoxXPath.TabStop = false;
            groupBoxXPath.Text = "XPath";
            // 
            // labelXPathExpression
            // 
            labelXPathExpression.AutoSize = true;
            labelXPathExpression.Location = new System.Drawing.Point(25, 22);
            labelXPathExpression.Name = "labelXPathExpression";
            labelXPathExpression.Size = new System.Drawing.Size(61, 13);
            labelXPathExpression.TabIndex = 3;
            labelXPathExpression.Text = "Expression:";
            // 
            // labelXPathPrefixMapping
            // 
            labelXPathPrefixMapping.AutoSize = true;
            labelXPathPrefixMapping.Location = new System.Drawing.Point(6, 59);
            labelXPathPrefixMapping.Name = "labelXPathPrefixMapping";
            labelXPathPrefixMapping.Size = new System.Drawing.Size(80, 13);
            labelXPathPrefixMapping.TabIndex = 4;
            labelXPathPrefixMapping.Text = "Prefix Mapping:";
            // 
            // textBoxXPathExpression
            // 
            this.textBoxXPathExpression.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxXPathExpression.Location = new System.Drawing.Point(92, 19);
            this.textBoxXPathExpression.Name = "textBoxXPathExpression";
            this.textBoxXPathExpression.ReadOnly = true;
            this.textBoxXPathExpression.Size = new System.Drawing.Size(302, 20);
            this.textBoxXPathExpression.TabIndex = 2;
            // 
            // textBoxXPathPrefixMapping
            // 
            this.textBoxXPathPrefixMapping.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxXPathPrefixMapping.Location = new System.Drawing.Point(92, 56);
            this.textBoxXPathPrefixMapping.Name = "textBoxXPathPrefixMapping";
            this.textBoxXPathPrefixMapping.ReadOnly = true;
            this.textBoxXPathPrefixMapping.Size = new System.Drawing.Size(302, 20);
            this.textBoxXPathPrefixMapping.TabIndex = 5;
            // 
            // XPathUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(groupBoxXPath);
            this.Name = "XPathUserControl";
            this.Size = new System.Drawing.Size(400, 82);
            groupBoxXPath.ResumeLayout(false);
            groupBoxXPath.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
    }
}