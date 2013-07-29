//------------------------------------------------------------------------------
// <copyright file="CustomXMLPartsForm.cs" company="Florian Wolters">
//     Copyright (c) Florian Wolters. All rights reserved.
// </copyright>
// <author>Florian Wolters &lt;wolters.fl@gmail.com&gt;</author>
//------------------------------------------------------------------------------

namespace FlorianWolters.Office.Word.AddIn.CBA.Forms
{
    using System.Collections.Generic;
    using System.Windows.Forms;
    using Office = Microsoft.Office.Core;

    internal partial class CustomXMLPartsForm : Form
    {
        public CustomXMLPartsForm(IEnumerable<Office.CustomXMLPart> customXMLParts)
            : this()
        {
            this.AddCustomXMLPartsToListView(customXMLParts);
        }

        public CustomXMLPartsForm()
        {
            this.InitializeComponent();
        }

        public string LastSelectedCustomXMLPartID { get; private set; }

        public void PopulateCustomXMLPartsListView(IEnumerable<Office.CustomXMLPart> customXMLParts)
        {
            this.ClearCustomXMLPartsListView();
            this.AddCustomXMLPartsToListView(customXMLParts);
        }

        private void ClearCustomXMLPartsListView()
        {
            this.listViewCustomXMLParts.Items.Clear();
        }

        private void AddCustomXMLPartsToListView(IEnumerable<Office.CustomXMLPart> customXMLParts)
        {
            foreach (Office.CustomXMLPart customXMLPart in customXMLParts)
            {
                this.AddCustomXMLPartToListView(customXMLPart);
            }
        }

        private void AddCustomXMLPartToListView(Office.CustomXMLPart customXMLPart)
        {
            ListViewItem item = this.listViewCustomXMLParts.Items.Add(customXMLPart.NamespaceURI);
            item.SubItems.AddRange(new[] { customXMLPart.Id });
            item.SubItems.AddRange(new[] { customXMLPart.XML });
        }

        private void SelectedIndexChanged_ListViewCustomXMLParts(object sender, System.EventArgs e)
        {
            ListView listView = (ListView)sender;

            this.buttonSelectCustomXMLPart.Enabled = 1 == listView.SelectedItems.Count;
        }

        private void OnClick_ButtonSelectCustomXMLPart(object sender, System.EventArgs e)
        {
            this.LastSelectedCustomXMLPartID = this.listViewCustomXMLParts.SelectedItems[0].SubItems[1].Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
