using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ImportClash
{
    public class XmlReportExtractor
    {
        public string ReportPath { get; private set; }
        public XmlReportExtractor(string reportPath)
        {
            this.ReportPath = reportPath;
        }

        public Intersection[] Extract()
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.Load(ReportPath);
            var clashes = xmlDoc.SelectNodes("//clashresult");
            Intersection[] clearedClash = new Intersection[clashes.Count];
            for (int i = 0; i < clashes.Count; i++)
            {
                XmlNode clash = clashes[i];
                var currentIntersection = new Intersection();
                currentIntersection.Name = clash.Attributes.GetNamedItem("name").Value;

                var childNodes = clash.ChildNodes;
                XmlNode clashObjects = null;
                XmlNode gridLocation = null;
                foreach (XmlNode child in childNodes)
                {
                    if (child.Name == "gridlocation") gridLocation = child;
                    if (child.Name == "clashobjects") clashObjects = child;
                }
                currentIntersection.GridLocation = gridLocation.InnerText;
                currentIntersection.InModelElement = new IntersectionElement(clashObjects.ChildNodes[0]);
                currentIntersection.InLinkElement = new IntersectionElement(clashObjects.ChildNodes[1]);
                clearedClash[i] = currentIntersection;
            }
            return clearedClash;
        }
    }

    public class ReportFinder
    {
        public string ModelDirectory { get; private set; }
        private string DialogIntialDirectory { get; set; }
        private bool ReportDirectoryFound { get; set; }

        public ReportFinder(string modelDirectory)
        {
            this.ModelDirectory = modelDirectory;
            this.DialogIntialDirectory = @"\\ois-revit1\Revit\Все проекты\";
            this.ReportDirectoryFound = false;
            //Found projecet related report folder
            var pathCutter = new Regex(@"[\s\|\S]*Хранилище отделов");
            var pathToCentral = pathCutter.Match(ModelDirectory).Value;
            if (pathToCentral != null)
            {
                var pathToClash = pathToCentral.Replace("Хранилище отделов", @"Сводная модель\Пересечения");
                if (Directory.Exists(pathToClash))
                {
                    var directories = Directory.GetDirectories(pathToClash);
                    var lastCreatedFolder = directories.Select(dirInfo => new FileInfo(dirInfo)).OrderByDescending(dirInfo => dirInfo.CreationTime).FirstOrDefault();
                    if (lastCreatedFolder != null) { DialogIntialDirectory = lastCreatedFolder.FullName; ReportDirectoryFound = true; }
                }
            }
        }

        public FileInfo TryFindAutomaticaly()
        {
            var reportMatchFound = false;
            string matchedReportName = null;
            if (ReportDirectoryFound)
            {
                var modelInDepartmentName = ModelDirectory.Split('\\').LastOrDefault();
                var files = Directory.GetFiles(DialogIntialDirectory, "*.html");
                foreach (var file in files)
                {
                    var inReportModelName = GetRelatedModelName(file);
                    if (inReportModelName != null && inReportModelName.Contains(modelInDepartmentName)) { matchedReportName = file; reportMatchFound = true; }
                }
            }
            if (reportMatchFound) return new FileInfo(matchedReportName);
            return this.FindWithDialog();
        }

        public FileInfo FindWithDialog()
        {
            using (var manualReportSelection = new OpenFileDialog())
            {
                manualReportSelection.InitialDirectory = this.DialogIntialDirectory;
                manualReportSelection.Filter = "Html(*.html)|*.html|Xml(*.xml)|*.xml";
                if (manualReportSelection.ShowDialog() == DialogResult.OK)
                {
                    return new FileInfo(manualReportSelection.FileName);
                }
            }
            return null;
        }

        private string GetRelatedModelName(string reportName)
        {
            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(reportName, true);
            string pattern = "//span[span='Элемент Файл источника']";
            var clashes = htmlDoc.DocumentNode.SelectSingleNode(pattern);
            if (clashes != null) { return (clashes.ChildNodes[1].InnerText); }
            return null;
        }


    }

    public class HtmlReportExtractor
    {
        public string ReportPath { get; private set; }
        public HtmlReportExtractor(string reportPath)
        {
            this.ReportPath = reportPath;
        }

        public Intersection[] Extract()
        {
            HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.Load(ReportPath, true);
            var clashes = htmlDoc.DocumentNode.SelectNodes("//div[@class='viewpoint']");
            Intersection[] clearedClash = new Intersection[clashes.Count];
            for (int i = 0; i < clashes.Count; i++)
            {
                var currentClash = clashes[i];
                var currentIntersection = new Intersection();
                var inModelElement = currentIntersection.InModelElement = new IntersectionElement();
                var inLinkElement = currentIntersection.InLinkElement = new IntersectionElement();
                IntersectionElement currentElement = null;
                bool notMeetIntersectionElement = true;

                foreach (var clild in currentClash.ChildNodes)
                {
                    var currentChilds = clild.ChildNodes;
                    if (currentChilds.Count == 2)
                    {
                        if (currentChilds[0].InnerText == "Имя") { currentIntersection.Name = currentChilds[1].InnerText; };
                        if (currentChilds[0].InnerText == "Расположение сетки") { currentIntersection.GridLocation = currentChilds[1].InnerText; };
                        if (currentChilds[0].InnerText == "Расположение сетки") { currentIntersection.GridLocation = currentChilds[1].InnerText; };
                        if (currentChilds[0].InnerText == "Объект Id") currentElement.ObjectId = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Элемент Файл источника") currentElement.ModelName = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Элемент Тип") currentElement.ElementType = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Объект Имя системы") currentElement.SystemName = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Объект Категория") currentElement.Category = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Объект Семейство") currentElement.Family = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Объект Тип") currentElement.ObjectType = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Объект Имя") currentElement.Name = currentChilds[1].InnerText;
                        if (currentChilds[0].InnerText == "Объект Рабочий набор") currentElement.Workset = currentChilds[1].InnerText;
                    }
                    else
                    {
                        if (clild.OuterHtml.Contains("href")) continue;
                        if (notMeetIntersectionElement) { notMeetIntersectionElement = false; currentElement = currentIntersection.InModelElement; }
                        else { currentElement = currentIntersection.InLinkElement; }
                    }
                }
                clearedClash[i] = currentIntersection;
            }
            return clearedClash;
        }
    }

}
