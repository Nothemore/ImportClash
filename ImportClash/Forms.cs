using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.Drawing;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace ImportClash
{



    public class UIForm : System.Windows.Forms.Form
    {
        private DataGridView dGrid;
        public TableLayoutPanel table;
        private PropertyInfo[] IntersectionElementProperies { get; set; }
        private int propertyCount;
        private Label conflictName;
        private Label gridLocation;
        public Label reportName;
        public Label reportDate;
        private CheckBox cutViewOnClashChange;
        public IntersectionsModel ClashModel { get; private set; }
        private TableLayoutPanel controlStore;
        private TableLayoutPanel labelStore;
        private OperationHandler ExHandler { get; set; }
        private ExternalEvent ExEvent { get; set; }
        public bool IsFormOpen { get; set; }
        private Button nextClash;
        private Button previousClash;
        private Button cutView;


        private UIForm()
        { }

        public UIForm(IntersectionsModel clashModel, OperationHandler exHandler, ExternalEvent exEvent, ExternalCommandData commandData)
        {
            ExEvent = exEvent;
            ExHandler = exHandler;
            ClashModel = clashModel;
            IntersectionElementProperies = typeof(IntersectionElement).GetProperties();
            propertyCount = IntersectionElementProperies.Length;
            ClashModel.StateChange += (args) => { UpdateIntersectionInfo(ClashModel.CurrentClash); };
            this.MaximizeBox = false;
            this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            //Метки
            reportDate = new Label() { Dock = DockStyle.Fill, Text = "Дата создания отчета111" };
            reportName = new Label() { Dock = DockStyle.Fill, Text = "Имя отчета" };
            conflictName = new Label() { Dock = DockStyle.Fill, Text = "Конфликт" };
            gridLocation = new Label() { Dock = DockStyle.Fill, Text = "Расположнение сетки" };

            labelStore = new TableLayoutPanel() { Dock = DockStyle.Fill };
            labelStore.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            labelStore.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            labelStore.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            labelStore.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            labelStore.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            labelStore.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            labelStore.CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset;

            labelStore.Controls.Add(reportDate, 0, 0);
            labelStore.Controls.Add(reportName, 1, 0);
            labelStore.Controls.Add(conflictName, 0, 1);
            labelStore.Controls.Add(gridLocation, 1, 1);


            //Управляющие контролы
            cutView = new Button() { Dock = DockStyle.Fill, Text = "Обрезать вид" };
            cutView.Click += (sender, args) =>
            {
                var Info = new ElementId(int.Parse(ClashModel.CurrentClash.InModelElement.ObjectId));
                this.DisableControls();
                ExHandler.Info.SetInfo(Info);
                ExEvent.Raise();
            };

            nextClash = new Button() { Dock = DockStyle.Fill, Text = "Следующее пересечение" };
            nextClash.Click += (sender, args) => { ClashModel.NextClash(); if (cutViewOnClashChange.Checked) cutView.PerformClick(); };
            previousClash = new Button() { Dock = DockStyle.Fill, Text = "Предыдущее пересечение" };
            previousClash.Click += (sender, args) => { ClashModel.PreviousClash(); if(cutViewOnClashChange.Checked) cutView.PerformClick(); };
            cutViewOnClashChange = new CheckBox() { Dock = DockStyle.Fill, Text = "Обрезать автоматически", Checked = false };
            Button opentReport = new Button() { Dock = DockStyle.Fill, Text = "Открыть отчет" };
            opentReport.Click += (sender, args) => 
            {
                
                var newReportRelatedClashModel = new IntersectionsModel(commandData.Application.ActiveUIDocument, false);
                if (!newReportRelatedClashModel.IsValid) { return ; };
                this.ClashModel = newReportRelatedClashModel;
                UpdateIntersectionInfo(ClashModel.CurrentClash);
                UpdateReportInfo();
                ClashModel.StateChange += (args1) => { UpdateIntersectionInfo(ClashModel.CurrentClash); };
            };

            //Для хранения управляющих элементов
            controlStore = new TableLayoutPanel() { Dock = DockStyle.Fill };
            for (var i = 0; i < 5; i++) controlStore.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            controlStore.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            controlStore.RowStyles.Add(new RowStyle(SizeType.Absolute, 35));
            controlStore.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            controlStore.Controls.Add(previousClash, 0, 0);
            controlStore.Controls.Add(nextClash, 1, 0);
            controlStore.Controls.Add(cutView, 2, 0);
            controlStore.Controls.Add(cutViewOnClashChange, 3, 0);
            controlStore.Controls.Add(opentReport, 4, 0);
            controlStore.Controls.Add(new System.Windows.Forms.Panel(),5,0);
            controlStore.CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset;

            // Датагрид
            dGrid = new DataGridView();
            dGrid.Dock = DockStyle.Fill;
            dGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dGrid.AutoGenerateColumns = false;
            dGrid.ReadOnly = true;
            dGrid.ColumnCount = 3;
            dGrid.Columns[0].Name = "Параметр";
            dGrid.Columns[1].Name = "Элемент в модели";
            dGrid.Columns[2].Name = "Элемент в связи";

            //Одна строка уже есть
            dGrid.AllowUserToAddRows = false;
            dGrid.AllowUserToOrderColumns = false;
            dGrid.RowHeadersVisible = false;
            dGrid.AllowUserToResizeRows = false;
            dGrid.AllowUserToResizeColumns = false;
         
            dGrid.ScrollBars = ScrollBars.None;
            for (int i = 0; i < propertyCount; i++) dGrid.Rows.Add();
            string[] parameterNames = new string[] { "Id", "Файла Источника", "Элемент Тип", "Имя системы",
           "Категория","Семейство","Объект Тип", "Объект Имя","Рабочий набор"};
            for (int i = 0; i < parameterNames.Length; i++) dGrid[0, i].Value = parameterNames[i];

            //Контейнер
            table = new TableLayoutPanel() { Dock = DockStyle.Fill };
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, labelStore.PreferredSize.Height + 8));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, controlStore.PreferredSize.Height + 8));
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            // table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.CellBorderStyle = TableLayoutPanelCellBorderStyle.Inset;

            //Добавление обеъктов в контейнер
            table.Controls.Add(labelStore, 0, 0);
            table.Controls.Add(controlStore, 0, 1);
            table.Controls.Add(dGrid, 0, 2);

            this.Controls.Add(table);

            commandData.Application.ApplicationClosing += (sender, args) => { this.Close(); };

            //Инициализация начального состояния
            this.Load += (sender, args) => { UpdateIntersectionInfo(ClashModel.CurrentClash); UpdateReportInfo(); };
        }

        private void UpdateIntersectionElementInfo(int columnNumber, IntersectionElement clashElement)
        {
            for (int i = 0; i < propertyCount; i++)
            {
                dGrid[columnNumber, i].Value = IntersectionElementProperies[i].GetValue(clashElement);
            }
        }

        private void UpdateReportInfo()
        {
            reportDate.Text = ClashModel.ReportInfo.CreationTimeUtc.ToString();
            reportName.Text = ClashModel.ReportInfo.FullName;
        }

        private void UpdateIntersectionInfo(Intersection currentIntersection)
        {
            conflictName.Text = currentIntersection.Name + $"   ({ClashModel.CurrentClashNumber + 1}/{ClashModel.TotalClash})";
            gridLocation.Text = "Расположение сетки " + currentIntersection.GridLocation;
            UpdateIntersectionElementInfo(1, currentIntersection.InModelElement);
            UpdateIntersectionElementInfo(2, currentIntersection.InLinkElement);
            var prefGrid = dGrid.PreferredSize;
            this.ClientSize = new Size(prefGrid.Width - 3, controlStore.PreferredSize.Height + labelStore.PreferredSize.Height + prefGrid.Height + 16);

        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            ExCommand.FormToShow.IsFormOpen = false;
            ExCommand.FormToShow = null;
            ExEvent.Dispose();
            ExEvent = null;
            ExHandler = null;
            base.OnFormClosed(e);
        }

        public void EnableControls()
        {
            foreach (System.Windows.Forms.Control ctrl in controlStore.Controls) ctrl.Enabled = true;
        }

        public void DisableControls()
        {
            foreach (System.Windows.Forms.Control ctrl in controlStore.Controls) ctrl.Enabled = false;
        }

    }



}
