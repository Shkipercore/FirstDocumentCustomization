using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;


namespace FirstDocumentCustomization
{
    public partial class FormEditWork : Form
    {
        public FormEditWork()
        {
            InitializeComponent();
        }

        private void buttonAddTypeWork_Click(object sender, EventArgs e)
        {
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;
            var myFactory = Globals.Ribbons.Ribbon1.Factory;

            if (textBoxAddTypeWork.Text.Length > 0)
            {
                RibbonDropDownItem insertItem = myFactory.CreateRibbonDropDownItem();
                insertItem.Label = textBoxAddTypeWork.Text;
                bool isItemNotPresent = true;
                foreach (var item in ribbon.comboBoxSelectionWork.Items)
                {
                    if (item.Label.Equals(insertItem.Label))
                        isItemNotPresent = false;
                }
                if (isItemNotPresent)
                {
                    checkedListBoxTypeWork.Items.Add(textBoxAddTypeWork.Text);
                    ribbon.comboBoxSelectionWork.Items.Add(insertItem);

                    EditorXML editorXML = new EditorXML();
                    editorXML.CreateNode(textBoxAddTypeWork.Text);
                }
            }
        }

        private void buttonDeleteWork_Click(object sender, EventArgs e)
        {
            EditorXML editorXML = new EditorXML();
            Ribbon1 ribbon = Globals.Ribbons.Ribbon1;

            var itemLength = ribbon.comboBoxSelectionWork.Items.Count;
            var SelectedItem = checkedListBoxTypeWork.SelectedItem;

            List<RibbonDropDownItem> listForRemove = new List<RibbonDropDownItem>();
            List<object> listForCheckBox = new List<object>();

            foreach (var item in ribbon.comboBoxSelectionWork.Items)
            {
                if (item.Label.Equals(SelectedItem))
                {
                    listForRemove.Add(item);
                }

            }

            foreach (var itemForRemove in listForRemove)
            {
                ribbon.comboBoxSelectionWork.Items.Remove(itemForRemove);
            }

            foreach (var item in checkedListBoxTypeWork.Items)
            {
                if (item == SelectedItem)
                {
                    listForCheckBox.Add(item);
                } 

            }

            foreach (var itemForCheckBox in listForCheckBox)
            {
                checkedListBoxTypeWork.Items.Remove(itemForCheckBox);
            }



            //for (int i = 0; i < itemLength; i++)
            //{
            //    var item = ribbon.comboBoxSelectionWork.Items[i];
            //    if (item.Label.Equals(SelectedItem))
            //    {
            //        ribbon.comboBoxSelectionWork.Items.Remove(item);
            //        checkedListBoxTypeWork.Items.Remove(SelectedItem);
            //        editorXML.RemoveElement(SelectedItem.ToString());
            //    }
            //}
        }
    }
}
