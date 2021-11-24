#region Namespaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

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
            return Result.Succeeded;
        }
    }
}