﻿//------------------------------------------------------------------------------
// <copyright file="XMLStructureUserControl.Designer.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
// <auto-generated/>
//------------------------------------------------------------------------------

namespace FlorianWolters.Windows.Forms.XML.UserControls
{
    partial class XMLStructureUserControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ImageList imageList;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.ToolStripButton toolStripButtonExpandNode;
        private System.Windows.Forms.ToolStripButton toolStripButtonCollapseNode;
        private System.Windows.Forms.ToolStripButton toolStripButtonSelectNode;
        internal System.Windows.Forms.TreeView treeViewStructure;

        /// <summary>
        /// Releases the unmanaged resources used by this <see cref="XPathUserControl"/> and optionally releases the
        /// managed resources.
        /// </summary>
        /// <param name="disposing">
        /// <c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged
        /// resources.
        /// </param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.ToolStripLabel toolStripLabel;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(XMLStructureUserControl));
            this.treeViewStructure = new System.Windows.Forms.TreeView();
            this.imageList = new System.Windows.Forms.ImageList(this.components);
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonExpandNode = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonCollapseNode = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonSelectNode = new System.Windows.Forms.ToolStripButton();
            toolStripLabel = new System.Windows.Forms.ToolStripLabel();
            this.toolStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripLabel
            // 
            toolStripLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            toolStripLabel.Name = "toolStripLabel";
            toolStripLabel.Size = new System.Drawing.Size(89, 22);
            toolStripLabel.Text = "XML Structure";
            // 
            // treeViewStructure
            // 
            this.treeViewStructure.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeViewStructure.Enabled = false;
            this.treeViewStructure.ImageIndex = 0;
            this.treeViewStructure.ImageList = this.imageList;
            this.treeViewStructure.Location = new System.Drawing.Point(0, 28);
            this.treeViewStructure.Name = "treeViewStructure";
            this.treeViewStructure.SelectedImageIndex = 0;
            this.treeViewStructure.Size = new System.Drawing.Size(320, 240);
            this.treeViewStructure.TabIndex = 0;
            this.treeViewStructure.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this.OnAfterCollapse);
            this.treeViewStructure.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.OnAfterExpand);
            this.treeViewStructure.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.OnAfterSelect);
            this.treeViewStructure.Enter += new System.EventHandler(this.OnEnterTreeViewStructure);
            this.treeViewStructure.Leave += new System.EventHandler(this.OnLeaveTreeViewStructure);
            // 
            // imageList
            // 
            this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
            this.imageList.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList.Images.SetKeyName(0, "xmlAttribute");
            this.imageList.Images.SetKeyName(1, "xmlComment");
            this.imageList.Images.SetKeyName(2, "xmlElementData");
            this.imageList.Images.SetKeyName(3, "xmlElementCollapsed");
            this.imageList.Images.SetKeyName(4, "xmlElementExpanded");
            // 
            // toolStrip
            // 
            this.toolStrip.Enabled = false;
            this.toolStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            toolStripLabel,
            this.toolStripButtonExpandNode,
            this.toolStripButtonCollapseNode,
            this.toolStripButtonSelectNode});
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.toolStrip.Size = new System.Drawing.Size(320, 25);
            this.toolStrip.TabIndex = 7;
            this.toolStrip.Text = "toolStrip";
            // 
            // toolStripButtonExpandNode
            // 
            this.toolStripButtonExpandNode.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonExpandNode.Image = global::FlorianWolters.Windows.Forms.XML.Properties.Resources.plus;
            this.toolStripButtonExpandNode.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonExpandNode.Name = "toolStripButtonExpandNode";
            this.toolStripButtonExpandNode.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonExpandNode.Text = "Expand One Level";
            this.toolStripButtonExpandNode.Click += new System.EventHandler(this.OnClickToolStripButtonExpandNode);
            // 
            // toolStripButtonCollapseNode
            // 
            this.toolStripButtonCollapseNode.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonCollapseNode.Image = global::FlorianWolters.Windows.Forms.XML.Properties.Resources.minus;
            this.toolStripButtonCollapseNode.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonCollapseNode.Name = "toolStripButtonCollapseNode";
            this.toolStripButtonCollapseNode.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonCollapseNode.Text = "Collapse One Level";
            this.toolStripButtonCollapseNode.Click += new System.EventHandler(this.OnClickToolStripButtonCollapseNode);
            // 
            // toolStripButtonSelectNode
            // 
            this.toolStripButtonSelectNode.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonSelectNode.Enabled = false;
            this.toolStripButtonSelectNode.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonSelectNode.Name = "toolStripButtonSelectNode";
            this.toolStripButtonSelectNode.Size = new System.Drawing.Size(74, 22);
            this.toolStripButtonSelectNode.Text = "Select Node";
            this.toolStripButtonSelectNode.ToolTipText = "Select the highlighted node";
            this.toolStripButtonSelectNode.Click += new System.EventHandler(this.OnClickToolStripButtonSelectNode);
            // 
            // XMLStructureUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.toolStrip);
            this.Controls.Add(this.treeViewStructure);
            this.Name = "XMLStructureUserControl";
            this.Size = new System.Drawing.Size(320, 268);
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
    }
}