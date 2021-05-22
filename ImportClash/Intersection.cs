using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.IO;

namespace ImportClash
{
    public class Intersection
    {
        public string Name { get; set; }
        public string GridLocation { get; set; }
        public IntersectionElement InModelElement { get; set; }
        public IntersectionElement InLinkElement { get; set; }
    }

    public class IntersectionElement
    {
        public string ObjectId { get; set; }
        public string ModelName { get; set; }
        public string ElementType { get; set; }
        public string SystemName { get; set; }
        public string Category { get; set; }
        public string Family { get; set; }
        public string ObjectType { get; set; }
        public string Name { get; set; }
        public string Workset { get; set; }

        public IntersectionElement() { }

        public IntersectionElement(XmlNode clashElement)
        {
            if (clashElement != null)
            {
                //Ищем smarttag среди дочерних нодов
                foreach (XmlNode nod in clashElement.ChildNodes)
                {
                    if (nod.ChildNodes.Count >= 9)
                    {
                        //Извлекаем информацию из нодов smarttag
                        foreach (XmlNode smartTag in nod.ChildNodes)
                        {
                            var childSmartTag = smartTag.ChildNodes;
                            if (childSmartTag[0].InnerText == "Объект Id") ObjectId = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Элемент Файл источника") ModelName = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Элемент Тип") ElementType = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Объект Имя системы") SystemName = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Объект Категория") Category = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Объект Семейство") Family = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Объект Тип") ObjectType = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Объект Имя") Name = childSmartTag[1].InnerText;
                            if (childSmartTag[0].InnerText == "Объект Рабочий набор") Workset = childSmartTag[1].InnerText;

                        }
                    }
                }
            }
        }
    }

    public class IntersectionsModel
    {
        public FileInfo ReportInfo { get; private set; }
        public Intersection[] ReportData { get; private set; }
        public bool IsValid { get; private set; }
        public Intersection CurrentClash { get; private set; }
        public int CurrentClashNumber { get; private set; }
        public int TotalClash { get; private set; }
        public event Action<Intersection> StateChange;

        public IntersectionsModel(UIDocument uiDoc, bool findReportAutomaticaly)
        {
            var centralPath = uiDoc.Document.GetWorksharingCentralModelPath();
            var userVisiblePath = ModelPathUtils.ConvertModelPathToUserVisiblePath(centralPath);
            var reportFinder = new ReportFinder(userVisiblePath);
            ReportInfo = findReportAutomaticaly ? reportFinder.TryFindAutomaticaly() : reportFinder.FindWithDialog();
            //if (ReportInfo == null) { throw new Exception("Не выбран файл отчета"); }
            if (ReportInfo == null) { TaskDialog.Show("Ошибка", "Не выбран файл отчета"); return; }
            Intersection[] reportData = null;
            var htmlExtractor = new HtmlReportExtractor(ReportInfo.FullName);
            reportData = htmlExtractor.Extract();
            //if (reportData == null) { throw new Exception("Не удалось распознать файл отчета"); }
            if (reportData == null) { TaskDialog.Show("Ошибка", "Не удалось распознать файл отчета"); return; };
            if (reportData != null && reportData.Length > 0)
            {
                ReportData = reportData;
                TotalClash = reportData.Length;
                UpdateCurrent(CurrentClashNumber = 0);
            }
            //else throw new Exception("Отчет о пересечениях не содержит элементов");
            else { TaskDialog.Show("Ошибка", "Отчет о пересечениях не содержит элементов"); return; };
            IsValid = true;
        }

        //Удалить
        private IntersectionsModel(Intersection[] reportData)
        {
            if (reportData != null && reportData.Length > 0)
            {
                ReportData = reportData;
                TotalClash = reportData.Length;
                UpdateCurrent(CurrentClashNumber = 0);

            }
            else throw new Exception("Отчет о пересечениях не содержит элементов");
        }

        public void NextClash()
        {
            if (CurrentClashNumber < ReportData.Length - 1)
            {
                UpdateCurrent(++CurrentClashNumber);
            }
        }

        public void PreviousClash()
        {
            if (CurrentClashNumber > 0)
            {
                UpdateCurrent(--CurrentClashNumber);
            }
        }

        private void UpdateCurrent(int currentClashNumber)
        {
            CurrentClash = ReportData[currentClashNumber];
            if (StateChange != null) StateChange(CurrentClash);
        }

    }
}
