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
            this.Font = new System.Windows.Forms.Label();
            this.sizeFont = new System.Windows.Forms.Label();
            this.colorFont = new System.Windows.Forms.Label();
            this.comboBoxSelectionFont = new System.Windows.Forms.ComboBox();
            this.comboBoxSelectionSizeFont = new System.Windows.Forms.ComboBox();
            this.buttonSaveSettings = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Font
            // 
            this.Font.AutoSize = true;
            this.Font.Location = new System.Drawing.Point(34, 41);
            this.Font.Name = "Font";
            this.Font.Size = new System.Drawing.Size(41, 13);
            this.Font.TabIndex = 0;
            this.Font.Text = "Шрифт";
            // 
            // sizeFont
            // 
            this.sizeFont.AutoSize = true;
            this.sizeFont.Location = new System.Drawing.Point(34, 68);
            this.sizeFont.Name = "sizeFont";
            this.sizeFont.Size = new System.Drawing.Size(88, 13);
            this.sizeFont.TabIndex = 1;
            this.sizeFont.Text = "Размер шрифта";
            // 
            // colorFont
            // 
            this.colorFont.AutoSize = true;
            this.colorFont.Location = new System.Drawing.Point(34, 94);
            this.colorFont.Name = "colorFont";
            this.colorFont.Size = new System.Drawing.Size(74, 13);
            this.colorFont.TabIndex = 2;
            this.colorFont.Text = "Цвет шрифта";
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
            // Settings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.buttonSaveSettings);
            this.Controls.Add(this.comboBoxSelectionSizeFont);
            this.Controls.Add(this.comboBoxSelectionFont);
            this.Controls.Add(this.colorFont);
            this.Controls.Add(this.sizeFont);
            this.Controls.Add(this.Font);
            this.Name = "Settings";
            this.Text = "Настройки";
            this.Load += new System.EventHandler(this.Settings_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Font;
        private System.Windows.Forms.Label sizeFont;
        private System.Windows.Forms.Label colorFont;
        private System.Windows.Forms.ComboBox comboBoxSelectionFont;
        private System.Windows.Forms.ComboBox comboBoxSelectionSizeFont;
        private System.Windows.Forms.Button buttonSaveSettings;
    }
}