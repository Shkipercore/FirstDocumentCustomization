namespace FirstDocumentCustomization
{
    partial class Settings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelFont = new System.Windows.Forms.Label();
            this.labelSizeFont = new System.Windows.Forms.Label();
            this.comboBoxSelectionFont = new System.Windows.Forms.ComboBox();
            this.comboBoxSelectionSizeFont = new System.Windows.Forms.ComboBox();
            this.buttonSaveSettings = new System.Windows.Forms.Button();
            this.labelColorFont = new System.Windows.Forms.Label();
            this.labelAlignmentText = new System.Windows.Forms.Label();
            this.comboBoxSelectionAlignmentText = new System.Windows.Forms.ComboBox();
            this.labelLineSpacing = new System.Windows.Forms.Label();
            this.labelLeftIndent = new System.Windows.Forms.Label();
            this.groupBoxParagraph = new System.Windows.Forms.GroupBox();
            this.labelRightIndent = new System.Windows.Forms.Label();
            this.labelIntervalBefore = new System.Windows.Forms.Label();
            this.labelIntervalAfter = new System.Windows.Forms.Label();
            this.groupBoxHeader = new System.Windows.Forms.GroupBox();
            this.comboBoxSelectionFontHeader = new System.Windows.Forms.ComboBox();
            this.labelFontHeader = new System.Windows.Forms.Label();
            this.labelSizeFontHeader = new System.Windows.Forms.Label();
            this.comboBoxSelectionSizeFontHeader = new System.Windows.Forms.ComboBox();
            this.labelAlignmentTextHeader = new System.Windows.Forms.Label();
            this.comboBoxSelectionAlignmentHeader = new System.Windows.Forms.ComboBox();
            this.groupBoxParagraph.SuspendLayout();
            this.groupBoxHeader.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelFont
            // 
            this.labelFont.AutoSize = true;
            this.labelFont.Location = new System.Drawing.Point(34, 41);
            this.labelFont.Name = "labelFont";
            this.labelFont.Size = new System.Drawing.Size(41, 13);
            this.labelFont.TabIndex = 0;
            this.labelFont.Text = "Шрифт";
            // 
            // labelSizeFont
            // 
            this.labelSizeFont.AutoSize = true;
            this.labelSizeFont.Location = new System.Drawing.Point(34, 71);
            this.labelSizeFont.Name = "labelSizeFont";
            this.labelSizeFont.Size = new System.Drawing.Size(88, 13);
            this.labelSizeFont.TabIndex = 1;
            this.labelSizeFont.Text = "Размер шрифта";
            // 
            // comboBoxSelectionFont
            // 
            this.comboBoxSelectionFont.FormattingEnabled = true;
            this.comboBoxSelectionFont.Items.AddRange(new object[] {
            "Arial",
            "Arial Black",
            "Calibri",
            "Comic Sans MS",
            "Courier New",
            "Franklin Gothic Medium",
            "Georgia",
            "Impact",
            "Lucida Console",
            "Lucida Sans Unicode",
            "Microsoft Sans Serif",
            "Palatino Linotype",
            "Sylfaen",
            "Tahoma",
            "Times New Roman",
            "Trebuchet MS",
            "Verdana"});
            this.comboBoxSelectionFont.Location = new System.Drawing.Point(81, 38);
            this.comboBoxSelectionFont.Name = "comboBoxSelectionFont";
            this.comboBoxSelectionFont.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSelectionFont.TabIndex = 3;
            // 
            // comboBoxSelectionSizeFont
            // 
            this.comboBoxSelectionSizeFont.FormattingEnabled = true;
            this.comboBoxSelectionSizeFont.Items.AddRange(new object[] {
            "8",
            "9",
            "10",
            "11",
            "12",
            "14",
            "16",
            "18",
            "20",
            "22",
            "24",
            "26",
            "28"});
            this.comboBoxSelectionSizeFont.Location = new System.Drawing.Point(129, 68);
            this.comboBoxSelectionSizeFont.Name = "comboBoxSelectionSizeFont";
            this.comboBoxSelectionSizeFont.Size = new System.Drawing.Size(48, 21);
            this.comboBoxSelectionSizeFont.TabIndex = 4;
            // 
            // buttonSaveSettings
            // 
            this.buttonSaveSettings.Location = new System.Drawing.Point(669, 373);
            this.buttonSaveSettings.Name = "buttonSaveSettings";
            this.buttonSaveSettings.Size = new System.Drawing.Size(75, 23);
            this.buttonSaveSettings.TabIndex = 5;
            this.buttonSaveSettings.Text = "Сохранить";
            this.buttonSaveSettings.UseVisualStyleBackColor = true;
            this.buttonSaveSettings.Click += new System.EventHandler(this.buttonSaveSettings_Click);
            // 
            // labelColorFont
            // 
            this.labelColorFont.AutoSize = true;
            this.labelColorFont.Location = new System.Drawing.Point(34, 383);
            this.labelColorFont.Name = "labelColorFont";
            this.labelColorFont.Size = new System.Drawing.Size(74, 13);
            this.labelColorFont.TabIndex = 2;
            this.labelColorFont.Text = "Цвет шрифта";
            // 
            // labelAlignmentText
            // 
            this.labelAlignmentText.AutoSize = true;
            this.labelAlignmentText.Location = new System.Drawing.Point(34, 103);
            this.labelAlignmentText.Name = "labelAlignmentText";
            this.labelAlignmentText.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.labelAlignmentText.Size = new System.Drawing.Size(119, 13);
            this.labelAlignmentText.TabIndex = 6;
            this.labelAlignmentText.Text = "Выравнивание текста";
            // 
            // comboBoxSelectionAlignmentText
            // 
            this.comboBoxSelectionAlignmentText.FormattingEnabled = true;
            this.comboBoxSelectionAlignmentText.Items.AddRange(new object[] {
            "По левому краю",
            "По центру",
            "По правому краю",
            "По ширине"});
            this.comboBoxSelectionAlignmentText.Location = new System.Drawing.Point(159, 100);
            this.comboBoxSelectionAlignmentText.Name = "comboBoxSelectionAlignmentText";
            this.comboBoxSelectionAlignmentText.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSelectionAlignmentText.TabIndex = 7;
            // 
            // labelLineSpacing
            // 
            this.labelLineSpacing.AutoSize = true;
            this.labelLineSpacing.Location = new System.Drawing.Point(6, 68);
            this.labelLineSpacing.Name = "labelLineSpacing";
            this.labelLineSpacing.Size = new System.Drawing.Size(139, 13);
            this.labelLineSpacing.TabIndex = 8;
            this.labelLineSpacing.Text = "Междустрочный интервал";
            // 
            // labelLeftIndent
            // 
            this.labelLeftIndent.AutoSize = true;
            this.labelLeftIndent.Location = new System.Drawing.Point(6, 16);
            this.labelLeftIndent.Name = "labelLeftIndent";
            this.labelLeftIndent.Size = new System.Drawing.Size(75, 13);
            this.labelLeftIndent.TabIndex = 9;
            this.labelLeftIndent.Text = "Отступ слева";
            // 
            // groupBoxParagraph
            // 
            this.groupBoxParagraph.Controls.Add(this.labelIntervalAfter);
            this.groupBoxParagraph.Controls.Add(this.labelLineSpacing);
            this.groupBoxParagraph.Controls.Add(this.labelIntervalBefore);
            this.groupBoxParagraph.Controls.Add(this.labelRightIndent);
            this.groupBoxParagraph.Controls.Add(this.labelLeftIndent);
            this.groupBoxParagraph.Location = new System.Drawing.Point(37, 127);
            this.groupBoxParagraph.Name = "groupBoxParagraph";
            this.groupBoxParagraph.Size = new System.Drawing.Size(245, 124);
            this.groupBoxParagraph.TabIndex = 10;
            this.groupBoxParagraph.TabStop = false;
            this.groupBoxParagraph.Text = "Абзац";
            // 
            // labelRightIndent
            // 
            this.labelRightIndent.AutoSize = true;
            this.labelRightIndent.Location = new System.Drawing.Point(6, 29);
            this.labelRightIndent.Name = "labelRightIndent";
            this.labelRightIndent.Size = new System.Drawing.Size(81, 13);
            this.labelRightIndent.TabIndex = 11;
            this.labelRightIndent.Text = "Отступ справа";
            // 
            // labelIntervalBefore
            // 
            this.labelIntervalBefore.AutoSize = true;
            this.labelIntervalBefore.Location = new System.Drawing.Point(6, 42);
            this.labelIntervalBefore.Name = "labelIntervalBefore";
            this.labelIntervalBefore.Size = new System.Drawing.Size(89, 13);
            this.labelIntervalBefore.TabIndex = 11;
            this.labelIntervalBefore.Text = "Интервал перед";
            // 
            // labelIntervalAfter
            // 
            this.labelIntervalAfter.AutoSize = true;
            this.labelIntervalAfter.Location = new System.Drawing.Point(6, 55);
            this.labelIntervalAfter.Name = "labelIntervalAfter";
            this.labelIntervalAfter.Size = new System.Drawing.Size(89, 13);
            this.labelIntervalAfter.TabIndex = 11;
            this.labelIntervalAfter.Text = "Интервал после";
            // 
            // groupBoxHeader
            // 
            this.groupBoxHeader.Controls.Add(this.comboBoxSelectionAlignmentHeader);
            this.groupBoxHeader.Controls.Add(this.labelAlignmentTextHeader);
            this.groupBoxHeader.Controls.Add(this.comboBoxSelectionSizeFontHeader);
            this.groupBoxHeader.Controls.Add(this.labelSizeFontHeader);
            this.groupBoxHeader.Controls.Add(this.labelFontHeader);
            this.groupBoxHeader.Controls.Add(this.comboBoxSelectionFontHeader);
            this.groupBoxHeader.Location = new System.Drawing.Point(329, 38);
            this.groupBoxHeader.Name = "groupBoxHeader";
            this.groupBoxHeader.Size = new System.Drawing.Size(255, 110);
            this.groupBoxHeader.TabIndex = 11;
            this.groupBoxHeader.TabStop = false;
            this.groupBoxHeader.Text = "Колонтитул";
            this.groupBoxHeader.Visible = false;
            // 
            // comboBoxSelectionFontHeader
            // 
            this.comboBoxSelectionFontHeader.FormattingEnabled = true;
            this.comboBoxSelectionFontHeader.Items.AddRange(new object[] {
            "Arial",
            "Arial Black",
            "Calibri",
            "Comic Sans MS",
            "Courier New",
            "Franklin Gothic Medium",
            "Georgia",
            "Impact",
            "Lucida Console",
            "Lucida Sans Unicode",
            "Microsoft Sans Serif",
            "Palatino Linotype",
            "Sylfaen",
            "Tahoma",
            "Times New Roman",
            "Trebuchet MS",
            "Verdana"});
            this.comboBoxSelectionFontHeader.Location = new System.Drawing.Point(53, 19);
            this.comboBoxSelectionFontHeader.Name = "comboBoxSelectionFontHeader";
            this.comboBoxSelectionFontHeader.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSelectionFontHeader.TabIndex = 4;
            // 
            // labelFontHeader
            // 
            this.labelFontHeader.AutoSize = true;
            this.labelFontHeader.Location = new System.Drawing.Point(6, 22);
            this.labelFontHeader.Name = "labelFontHeader";
            this.labelFontHeader.Size = new System.Drawing.Size(41, 13);
            this.labelFontHeader.TabIndex = 5;
            this.labelFontHeader.Text = "Шрифт";
            // 
            // labelSizeFontHeader
            // 
            this.labelSizeFontHeader.AutoSize = true;
            this.labelSizeFontHeader.Location = new System.Drawing.Point(6, 49);
            this.labelSizeFontHeader.Name = "labelSizeFontHeader";
            this.labelSizeFontHeader.Size = new System.Drawing.Size(88, 13);
            this.labelSizeFontHeader.TabIndex = 6;
            this.labelSizeFontHeader.Text = "Размер шрифта";
            // 
            // comboBoxSelectionSizeFontHeader
            // 
            this.comboBoxSelectionSizeFontHeader.FormattingEnabled = true;
            this.comboBoxSelectionSizeFontHeader.Items.AddRange(new object[] {
            "8",
            "9",
            "10",
            "11",
            "12",
            "14",
            "16",
            "18",
            "20",
            "22",
            "24",
            "26",
            "28"});
            this.comboBoxSelectionSizeFontHeader.Location = new System.Drawing.Point(100, 46);
            this.comboBoxSelectionSizeFontHeader.Name = "comboBoxSelectionSizeFontHeader";
            this.comboBoxSelectionSizeFontHeader.Size = new System.Drawing.Size(48, 21);
            this.comboBoxSelectionSizeFontHeader.TabIndex = 7;
            // 
            // labelAlignmentTextHeader
            // 
            this.labelAlignmentTextHeader.AutoSize = true;
            this.labelAlignmentTextHeader.Location = new System.Drawing.Point(6, 80);
            this.labelAlignmentTextHeader.Name = "labelAlignmentTextHeader";
            this.labelAlignmentTextHeader.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.labelAlignmentTextHeader.Size = new System.Drawing.Size(119, 13);
            this.labelAlignmentTextHeader.TabIndex = 8;
            this.labelAlignmentTextHeader.Text = "Выравнивание текста";
            // 
            // comboBoxSelectionAlignmentHeader
            // 
            this.comboBoxSelectionAlignmentHeader.FormattingEnabled = true;
            this.comboBoxSelectionAlignmentHeader.Items.AddRange(new object[] {
            "По левому краю",
            "По центру",
            "По правому краю",
            "По ширине"});
            this.comboBoxSelectionAlignmentHeader.Location = new System.Drawing.Point(128, 77);
            this.comboBoxSelectionAlignmentHeader.Name = "comboBoxSelectionAlignmentHeader";
            this.comboBoxSelectionAlignmentHeader.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSelectionAlignmentHeader.TabIndex = 9;
            // 
            // Settings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBoxHeader);
            this.Controls.Add(this.groupBoxParagraph);
            this.Controls.Add(this.comboBoxSelectionAlignmentText);
            this.Controls.Add(this.labelAlignmentText);
            this.Controls.Add(this.buttonSaveSettings);
            this.Controls.Add(this.comboBoxSelectionSizeFont);
            this.Controls.Add(this.comboBoxSelectionFont);
            this.Controls.Add(this.labelColorFont);
            this.Controls.Add(this.labelSizeFont);
            this.Controls.Add(this.labelFont);
            this.Name = "Settings";
            this.Text = "Настройки";
            this.Load += new System.EventHandler(this.Settings_Load);
            this.groupBoxParagraph.ResumeLayout(false);
            this.groupBoxParagraph.PerformLayout();
            this.groupBoxHeader.ResumeLayout(false);
            this.groupBoxHeader.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelFont;
        private System.Windows.Forms.Label labelSizeFont;
        private System.Windows.Forms.ComboBox comboBoxSelectionFont;
        private System.Windows.Forms.ComboBox comboBoxSelectionSizeFont;
        private System.Windows.Forms.Button buttonSaveSettings;
        private System.Windows.Forms.Label labelColorFont;
        private System.Windows.Forms.Label labelAlignmentText;
        private System.Windows.Forms.ComboBox comboBoxSelectionAlignmentText;
        private System.Windows.Forms.Label labelLineSpacing;
        private System.Windows.Forms.Label labelLeftIndent;
        private System.Windows.Forms.GroupBox groupBoxParagraph;
        private System.Windows.Forms.Label labelIntervalAfter;
        private System.Windows.Forms.Label labelIntervalBefore;
        private System.Windows.Forms.Label labelRightIndent;
        private System.Windows.Forms.GroupBox groupBoxHeader;
        private System.Windows.Forms.ComboBox comboBoxSelectionAlignmentHeader;
        private System.Windows.Forms.Label labelAlignmentTextHeader;
        private System.Windows.Forms.ComboBox comboBoxSelectionSizeFontHeader;
        private System.Windows.Forms.Label labelSizeFontHeader;
        private System.Windows.Forms.Label labelFontHeader;
        private System.Windows.Forms.ComboBox comboBoxSelectionFontHeader;
    }
}