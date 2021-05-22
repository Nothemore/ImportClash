using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ImportClash
{
    public class OperationHandler : IExternalEventHandler
    {
        public OperationInfo Info { get; private set; }

        public OperationHandler() { Info = new OperationInfo(); }


        public void Execute(UIApplication app)
        {
            setSelection(Info.GetInfo(), app.ActiveUIDocument);
        }

        public string GetName()
        {
            return "ImportClashHandler";
        }

        public void setSelection(ElementId elemId, UIDocument uiDoc)
        {
            ExCommand.FormToShow.EnableControls();
            if (uiDoc.Document.GetElement(elemId) == null) { TaskDialog.Show("Внимание", "Не удалось найти элемент в модели");return; }

            using (var tr = new Transaction(uiDoc.Document))
            {
                tr.Start("FocusOnClash");
                uiDoc.Selection.SetElementIds(new List<ElementId>() { elemId });
                uiDoc.Application.PostCommand(RevitCommandId.LookupPostableCommandId(PostableCommand.SelectionBox));
                tr.Commit();
            }
            
            
        }
    }

    public class OperationInfo
    {
        private ElementId elemId = ElementId.InvalidElementId;
        public ElementId ElementsId { get { return elemId; } }

        public OperationInfo()
        {
            elemId = new ElementId(ElementId.InvalidElementId.IntegerValue);
        }

        public void SetInfo(ElementId newElemId)
        {
            Interlocked.Exchange(ref this.elemId, newElemId);
        }

        public ElementId GetInfo()
        {
            return Interlocked.Exchange(ref this.elemId, ElementId.InvalidElementId);
        }
    }

    
}
