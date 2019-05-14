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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.comboBoxSelectionWork = this.Factory.CreateRibbonComboBox();
            this.buttonApply = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.buttonFont = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.editBoxLeftIndent = this.Factory.CreateRibbonEditBox();
            this.editBoxRightIndent = this.Factory.CreateRibbonEditBox();
            this.editBoxLineSpacing = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.editBoxIntervalBefore = this.Factory.CreateRibbonEditBox();
            this.editBoxIntervalAfter = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonSettings = this.Factory.CreateRibbonButton();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.comboBoxAlignmentText = this.Factory.CreateRibbonComboBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Проверка оформления работ";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.comboBoxSelectionWork);
            this.group1.Items.Add(this.buttonApply);
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
            this.comboBoxSelectionWork.Text = null;
            // 
            // buttonApply
            // 
            this.buttonApply.Label = "Применить";
            this.buttonApply.Name = "buttonApply";
            this.buttonApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonApply_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.buttonFont);
            this.group3.Items.Add(this.comboBoxAlignmentText);
            this.group3.Label = "Текст";
            this.group3.Name = "group3";
            // 
            // buttonFont
            // 
            this.buttonFont.Label = "Шрифт";
            this.buttonFont.Name = "buttonFont";
            this.buttonFont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFont_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.editBoxLeftIndent);
            this.group4.Items.Add(this.editBoxRightIndent);
            this.group4.Items.Add(this.editBoxLineSpacing);
            this.group4.Items.Add(this.separator1);
            this.group4.Items.Add(this.editBoxIntervalBefore);
            this.group4.Items.Add(this.editBoxIntervalAfter);
            this.group4.Label = "Абзац";
            this.group4.Name = "group4";
            // 
            // editBoxLeftIndent
            // 
            this.editBoxLeftIndent.Label = "Отступ слева  ";
            this.editBoxLeftIndent.Name = "editBoxLeftIndent";
            this.editBoxLeftIndent.Text = null;
            // 
            // editBoxRightIndent
            // 
            this.editBoxRightIndent.Label = "Отступ справа";
            this.editBoxRightIndent.Name = "editBoxRightIndent";
            this.editBoxRightIndent.Text = null;
            // 
            // editBoxLineSpacing
            // 
            this.editBoxLineSpacing.Label = "Междустрочный интервал";
            this.editBoxLineSpacing.Name = "editBoxLineSpacing";
            this.editBoxLineSpacing.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // editBoxIntervalBefore
            // 
            this.editBoxIntervalBefore.Label = "Интервал перед";
            this.editBoxIntervalBefore.Name = "editBoxIntervalBefore";
            this.editBoxIntervalBefore.Text = null;
            // 
            // editBoxIntervalAfter
            // 
            this.editBoxIntervalAfter.Label = "Интервал после";
            this.editBoxIntervalAfter.Name = "editBoxIntervalAfter";
            this.editBoxIntervalAfter.Text = null;
            // 
            // group2
            // 
            this.group2.Items.Add(this.buttonSettings);
            this.group2.Name = "group2";
            this.group2.Visible = false;
            // 
            // buttonSettings
            // 
            this.buttonSettings.Label = "Настройки";
            this.buttonSettings.Name = "buttonSettings";
            this.buttonSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSettings_Click);
            // 
            // comboBoxAlignmentText
            // 
            ribbonDropDownItemImpl4.Label = "По левому краю";
            ribbonDropDownItemImpl5.Label = "По центру";
            ribbonDropDownItemImpl6.Label = "По правому краю";
            ribbonDropDownItemImpl7.Label = "По ширине";
            this.comboBoxAlignmentText.Items.Add(ribbonDropDownItemImpl4);
            this.comboBoxAlignmentText.Items.Add(ribbonDropDownItemImpl5);
            this.comboBoxAlignmentText.Items.Add(ribbonDropDownItemImpl6);
            this.comboBoxAlignmentText.Items.Add(ribbonDropDownItemImpl7);
            this.comboBoxAlignmentText.Label = "Выравнивание текста";
            this.comboBoxAlignmentText.Name = "comboBoxAlignmentText";
            this.comboBoxAlignmentText.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxAlignmentText_TextChanged);
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
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFont;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.FontDialog fontDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxLeftIndent;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxRightIndent;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxIntervalBefore;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxIntervalAfter;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxLineSpacing;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxAlignmentText;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
