namespace FirstDocumentCustomization
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.comboBoxSelectionWork = this.Factory.CreateRibbonComboBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonSettings = this.Factory.CreateRibbonButton();
            this.buttonApply = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Проверка оформления работ";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.comboBoxSelectionWork);
            this.group1.Name = "group1";
            // 
            // comboBoxSelectionWork
            // 
            ribbonDropDownItemImpl1.Label = "Курсовая работа";
            ribbonDropDownItemImpl2.Label = "Лабораторная работа";
            ribbonDropDownItemImpl3.Label = "ВКР";
            this.comboBoxSelectionWork.Items.Add(ribbonDropDownItemImpl1);
            this.comboBoxSelectionWork.Items.Add(ribbonDropDownItemImpl2);
            this.comboBoxSelectionWork.Items.Add(ribbonDropDownItemImpl3);
            this.comboBoxSelectionWork.Label = "Тип работы";
            this.comboBoxSelectionWork.Name = "comboBoxSelectionWork";
            this.comboBoxSelectionWork.SizeString = "500000000000000000000";
            this.comboBoxSelectionWork.Text = null;
            // 
            // group2
            // 
            this.group2.Items.Add(this.buttonSettings);
            this.group2.Items.Add(this.buttonApply);
            this.group2.Name = "group2";
            // 
            // buttonSettings
            // 
            this.buttonSettings.Label = "Настройки";
            this.buttonSettings.Name = "buttonSettings";
            this.buttonSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSettings_Click);
            // 
            // buttonApply
            // 
            this.buttonApply.Label = "Применить";
            this.buttonApply.Name = "buttonApply";
            this.buttonApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonApply_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxSelectionWork;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonApply;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
