using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ImportClash
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class ExCommand : IExternalCommand
    {
        public static UIForm FormToShow;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            if (!commandData.Application.ActiveUIDocument.Document.IsWorkshared)
            {
                TaskDialog.Show("Ошибка", "Режим совместной работы выключен");
                return Result.Succeeded;
            }
            if (FormToShow == null)
            {
                var clashModel = new IntersectionsModel(commandData.Application.ActiveUIDocument, true);
                if (!clashModel.IsValid) { return Result.Succeeded; };
                OperationHandler exHandler = new OperationHandler();
                ExternalEvent exEvent = ExternalEvent.Create(exHandler);
                FormToShow = new UIForm(clashModel, exHandler, exEvent, commandData);
                FormToShow.Show();
            }
            return Result.Succeeded;
        }
    }
}
