namespace VisioAddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.fileGroup = this.Factory.CreateRibbonGroup();
            this.openFileButton = this.Factory.CreateRibbonButton();
            this.exportGraphButton = this.Factory.CreateRibbonButton();
            this.toolsGroup = this.Factory.CreateRibbonGroup();
            this.selectMenu = this.Factory.CreateRibbonMenu();
            this.selectAllNodesButton = this.Factory.CreateRibbonButton();
            this.selectConnectedNodeButton = this.Factory.CreateRibbonButton();
            this.selectNonConnectedNodesButton = this.Factory.CreateRibbonButton();
            this.selectEdgesButton = this.Factory.CreateRibbonButton();
            this.invertButton = this.Factory.CreateRibbonButton();
            this.layoutButton = this.Factory.CreateRibbonButton();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.tab1.SuspendLayout();
            this.fileGroup.SuspendLayout();
            this.toolsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.fileGroup);
            this.tab1.Groups.Add(this.toolsGroup);
            this.tab1.Label = "DOT";
            this.tab1.Name = "tab1";
            // 
            // fileGroup
            // 
            this.fileGroup.Items.Add(this.openFileButton);
            this.fileGroup.Items.Add(this.exportGraphButton);
            this.fileGroup.Label = "Файл";
            this.fileGroup.Name = "fileGroup";
            // 
            // openFileButton
            // 
            this.openFileButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.openFileButton.Description = "Выбрать файл";
            this.openFileButton.Image = global::VisioAddIn.Properties.Resources.Browse;
            this.openFileButton.ImageName = "Выбрать файл";
            this.openFileButton.Label = "Выбрать файл";
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.ShowImage = true;
            this.openFileButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openFileButton_Click);
            // 
            // exportGraphButton
            // 
            this.exportGraphButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.exportGraphButton.Image = global::VisioAddIn.Properties.Resources.Save;
            this.exportGraphButton.Label = "Сохранить изменения";
            this.exportGraphButton.Name = "exportGraphButton";
            this.exportGraphButton.ShowImage = true;
            this.exportGraphButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportGraphButton_Click);
            // 
            // toolsGroup
            // 
            this.toolsGroup.Items.Add(this.selectMenu);
            this.toolsGroup.Items.Add(this.invertButton);
            this.toolsGroup.Items.Add(this.layoutButton);
            this.toolsGroup.Label = "Инструменты";
            this.toolsGroup.Name = "toolsGroup";
            // 
            // selectMenu
            // 
            this.selectMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.selectMenu.Image = global::VisioAddIn.Properties.Resources.Select;
            this.selectMenu.Items.Add(this.selectAllNodesButton);
            this.selectMenu.Items.Add(this.selectConnectedNodeButton);
            this.selectMenu.Items.Add(this.selectNonConnectedNodesButton);
            this.selectMenu.Items.Add(this.selectEdgesButton);
            this.selectMenu.Label = "Выделить";
            this.selectMenu.Name = "selectMenu";
            this.selectMenu.ShowImage = true;
            // 
            // selectAllNodesButton
            // 
            this.selectAllNodesButton.Label = "Все вершины";
            this.selectAllNodesButton.Name = "selectAllNodesButton";
            this.selectAllNodesButton.ShowImage = true;
            this.selectAllNodesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectAllNodesButton_Click);
            // 
            // selectConnectedNodeButton
            // 
            this.selectConnectedNodeButton.Label = "Соединенные вершины";
            this.selectConnectedNodeButton.Name = "selectConnectedNodeButton";
            this.selectConnectedNodeButton.ShowImage = true;
            this.selectConnectedNodeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectConnectedNodeButton_Click);
            // 
            // selectNonConnectedNodesButton
            // 
            this.selectNonConnectedNodesButton.Label = "Несоединенные вершины";
            this.selectNonConnectedNodesButton.Name = "selectNonConnectedNodesButton";
            this.selectNonConnectedNodesButton.ShowImage = true;
            this.selectNonConnectedNodesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectNonConnectedNodesButton_Click);
            // 
            // selectEdgesButton
            // 
            this.selectEdgesButton.Label = "Все ребра";
            this.selectEdgesButton.Name = "selectEdgesButton";
            this.selectEdgesButton.ShowImage = true;
            this.selectEdgesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectEdgesButton_Click);
            // 
            // invertButton
            // 
            this.invertButton.Label = "Инвертировать ребро";
            this.invertButton.Name = "invertButton";
            this.invertButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.invertButton_Click);
            // 
            // layoutButton
            // 
            this.layoutButton.Label = "Планировка";
            this.layoutButton.Name = "layoutButton";
            this.layoutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.layoutButton_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "Graph File";
            this.openFileDialog.Filter = "DOT files (*.gv;*.dot)|*.gv;*.dot";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "gv";
            this.saveFileDialog.Filter = "DOT files (*.gv;*.dot)|*.gv;*.dot";
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.fileGroup.ResumeLayout(false);
            this.fileGroup.PerformLayout();
            this.toolsGroup.ResumeLayout(false);
            this.toolsGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup fileGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton openFileButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportGraphButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup toolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu selectMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectAllNodesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectConnectedNodeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectNonConnectedNodesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton invertButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton layoutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectEdgesButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
