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
            this.buttonEditWork = this.Factory.CreateRibbonButton();
            this.buttonApply = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.buttonFont = this.Factory.CreateRibbonButton();
            this.comboBoxAlignmentText = this.Factory.CreateRibbonComboBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.editBoxLeftIndent = this.Factory.CreateRibbonEditBox();
            this.editBoxRightIndent = this.Factory.CreateRibbonEditBox();
            this.editBoxFirstLineIndent = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.editBoxIntervalBefore = this.Factory.CreateRibbonEditBox();
            this.editBoxIntervalAfter = this.Factory.CreateRibbonEditBox();
            this.editBoxLineSpacing = this.Factory.CreateRibbonEditBox();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Проверка оформления работ";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.comboBoxSelectionWork);
            this.group1.Items.Add(this.buttonEditWork);
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
            this.comboBoxSelectionWork.MaxLength = 200;
            this.comboBoxSelectionWork.Name = "comboBoxSelectionWork";
            this.comboBoxSelectionWork.Text = null;
            this.comboBoxSelectionWork.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.comboBoxSelectionWork_TextChanged);
            // 
            // buttonEditWork
            // 
            this.buttonEditWork.Label = "Редактировать";
            this.buttonEditWork.Name = "buttonEditWork";
            this.buttonEditWork.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonEditWork_Click);
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
            this.buttonFont.OfficeImageId = "ContentControlRichText";
            this.buttonFont.ShowImage = true;
            this.buttonFont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFont_Click);
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
            this.comboBoxAlignmentText.OfficeImageId = "AlignCenter";
            this.comboBoxAlignmentText.ShowImage = true;
            this.comboBoxAlignmentText.Text = null;
            // 
            // group4
            // 
            this.group4.Items.Add(this.editBoxLeftIndent);
            this.group4.Items.Add(this.editBoxRightIndent);
            this.group4.Items.Add(this.editBoxFirstLineIndent);
            this.group4.Items.Add(this.separator1);
            this.group4.Items.Add(this.editBoxIntervalBefore);
            this.group4.Items.Add(this.editBoxIntervalAfter);
            this.group4.Items.Add(this.editBoxLineSpacing);
            this.group4.Label = "Абзац";
            this.group4.Name = "group4";
            // 
            // editBoxLeftIndent
            // 
            this.editBoxLeftIndent.Label = "Отступ слева  ";
            this.editBoxLeftIndent.Name = "editBoxLeftIndent";
            this.editBoxLeftIndent.OfficeImageId = "IndentClassic";
            this.editBoxLeftIndent.ShowImage = true;
            this.editBoxLeftIndent.Text = null;
            this.editBoxLeftIndent.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBoxLeftIndent_TextChanged);
            // 
            // editBoxRightIndent
            // 
            this.editBoxRightIndent.Label = "Отступ справа";
            this.editBoxRightIndent.Name = "editBoxRightIndent";
            this.editBoxRightIndent.OfficeImageId = "IndentRTL";
            this.editBoxRightIndent.ShowImage = true;
            this.editBoxRightIndent.Text = null;
            // 
            // editBoxFirstLineIndent
            // 
            this.editBoxFirstLineIndent.Label = "Отступ первой строки";
            this.editBoxFirstLineIndent.Name = "editBoxFirstLineIndent";
            this.editBoxFirstLineIndent.OfficeImageId = "AlignJustifyMedium";
            this.editBoxFirstLineIndent.ShowImage = true;
            this.editBoxFirstLineIndent.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // editBoxIntervalBefore
            // 
            this.editBoxIntervalBefore.Label = "Интервал перед";
            this.editBoxIntervalBefore.Name = "editBoxIntervalBefore";
            this.editBoxIntervalBefore.OfficeImageId = "ParagraphSpacingBefore";
            this.editBoxIntervalBefore.ShowImage = true;
            this.editBoxIntervalBefore.Text = null;
            // 
            // editBoxIntervalAfter
            // 
            this.editBoxIntervalAfter.Label = "Интервал после";
            this.editBoxIntervalAfter.Name = "editBoxIntervalAfter";
            this.editBoxIntervalAfter.OfficeImageId = "ParagraphSpacingAfter";
            this.editBoxIntervalAfter.ShowImage = true;
            this.editBoxIntervalAfter.Text = null;
            // 
            // editBoxLineSpacing
            // 
            this.editBoxLineSpacing.Label = "Междустрочный интервал";
            this.editBoxLineSpacing.Name = "editBoxLineSpacing";
            this.editBoxLineSpacing.OfficeImageId = "LineSpacing";
            this.editBoxLineSpacing.ShowImage = true;
            this.editBoxLineSpacing.Text = null;
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxSelectionWork;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxFirstLineIndent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonEditWork;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
