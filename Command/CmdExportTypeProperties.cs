#region Namespaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

#endregion Namespaces

namespace TepscoImportExport.Command
{
    [Transaction(TransactionMode.Manual)]
    public class CmdExportTypeProperties : IExternalCommand
    {
        private UIDocument _uiDoc;
        private Document _doc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            _uiDoc = uiapp.ActiveUIDocument;
            _doc = _uiDoc.Document;

            // Get list category
            if (_doc.IsFamilyDocument)
                return Result.Failed;

            List<string> lstCategory = new List<string>();

            lstCategory = Common.ListCategories(_uiDoc, _doc);

            return Result.Succeeded;
        }
    }
}