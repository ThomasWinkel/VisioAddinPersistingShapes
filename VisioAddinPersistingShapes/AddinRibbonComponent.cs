using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioAddinPersistingShapes
{
    public partial class AddinRibbonComponent
    {
        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.IVShape shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.PrimaryItem;
            if (shape is null) return;

            DataObject dataObj = new DataObject(shape);
            foreach (string clipFormat in dataObj.GetFormats(false))
            {
                if (clipFormat == ddFormat.SelectedItem.Label)
                {
                    MemoryStream stream = (MemoryStream)dataObj.GetData(clipFormat);
                    string shapeData = Convert.ToBase64String(stream.ToArray());
                    Clipboard.SetText(shapeData);
                }
            }
        }

        private void btnLoad_Click(object sender, RibbonControlEventArgs e)
        {
            string shapeData = Clipboard.GetText();
            MemoryStream stream = new MemoryStream(Convert.FromBase64String(shapeData));
            DataObject obj = new DataObject(ddFormat.SelectedItem.Label, stream);
            Globals.ThisAddIn.Application.ActivePage.Drop(obj, 0, 0);
        }

        private void btnListFormats_Click(object sender, RibbonControlEventArgs e)
        {
            Visio.IVShape shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.PrimaryItem;
            if (shape is null) return;

            ddFormat.Items.Clear();

            DataObject dataObj = new DataObject(shape);
            foreach (string clipFormat in dataObj.GetFormats(false))
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = clipFormat;
                ddFormat.Items.Add(item);
            }
        }
    }
}