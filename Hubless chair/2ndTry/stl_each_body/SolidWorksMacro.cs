using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public partial class SolidWorksMacro
{
    public void Main()
    {
        SldWorks swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
        ModelDoc2 swModel = swApp.ActiveDoc as ModelDoc2;
        if (swModel == null)
        {
            System.Windows.Forms.MessageBox.Show("No active document.");
            return;
        }

        if (swModel.GetType() != (int)swDocumentTypes_e.swDocPART)
        {
            System.Windows.Forms.MessageBox.Show("Active document is not a part.");
            return;
        }

        PartDoc swPart = (PartDoc)swModel;
        object[] vBodies = (object[])swPart.GetBodies2((int)swBodyType_e.swAllBodies, true);

        if (vBodies == null || vBodies.Length == 0)
        {
            System.Windows.Forms.MessageBox.Show("No bodies found in the part.");
            return;
        }

        string path = System.IO.Path.GetDirectoryName(swModel.GetPathName());
        if (string.IsNullOrEmpty(path))
        {
            System.Windows.Forms.MessageBox.Show("Please save the part document first.");
            return;
        }

        ExportStlData swExportStlData = (ExportStlData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportStl);
        for (int i = 0; i < vBodies.Length; i++)
        {
            Body2 swBody = (Body2)vBodies[i];
            string fileName = System.IO.Path.Combine(path, "Body_" + (i + 1) + ".stl");
            swExportStlData.SetBody(swBody);
            swModel.Extension.SaveAs(fileName, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, (int)swFileSaveTypes_e.swSTL, swExportStlData, null, null);
        }

        System.Windows.Forms.MessageBox.Show("Bodies saved as separate STL files.");
    }

    public SldWorks SwApp { get; set; }
}
