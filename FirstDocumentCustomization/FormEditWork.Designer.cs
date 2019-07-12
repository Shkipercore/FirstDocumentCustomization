namespace FirstDocumentCustomization
{
    partial class FormEditWork
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
            this.checkedListBoxTypeWork = new System.Windows.Forms.CheckedListBox();
            this.buttonAddTypeWork = new System.Windows.Forms.Button();
            this.buttonDeleteWork = new System.Windows.Forms.Button();
            this.textBoxAddTypeWork = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // checkedListBoxTypeWork
            // 
            this.checkedListBoxTypeWork.CheckOnClick = true;
            this.checkedListBoxTypeWork.FormattingEnabled = true;
            this.checkedListBoxTypeWork.Location = new System.Drawing.Point(28, 29);
            this.checkedListBoxTypeWork.Name = "checkedListBoxTypeWork";
            this.checkedListBoxTypeWork.Size = new System.Drawing.Size(180, 94);
            this.checkedListBoxTypeWork.TabIndex = 0;
            // 
            // buttonAddTypeWork
            // 
            this.buttonAddTypeWork.Location = new System.Drawing.Point(28, 129);
            this.buttonAddTypeWork.Name = "buttonAddTypeWork";
            this.buttonAddTypeWork.Size = new System.Drawing.Size(75, 23);
            this.buttonAddTypeWork.TabIndex = 2;
            this.buttonAddTypeWork.Text = "Добавить";
            this.buttonAddTypeWork.UseVisualStyleBackColor = true;
            this.buttonAddTypeWork.Click += new System.EventHandler(this.buttonAddTypeWork_Click);
            // 
            // buttonDeleteWork
            // 
            this.buttonDeleteWork.Location = new System.Drawing.Point(28, 159);
            this.buttonDeleteWork.Name = "buttonDeleteWork";
            this.buttonDeleteWork.Size = new System.Drawing.Size(75, 23);
            this.buttonDeleteWork.TabIndex = 3;
            this.buttonDeleteWork.Text = "Удалить";
            this.buttonDeleteWork.UseVisualStyleBackColor = true;
            this.buttonDeleteWork.Click += new System.EventHandler(this.buttonDeleteWork_Click);
            // 
            // textBoxAddTypeWork
            // 
            this.textBoxAddTypeWork.Location = new System.Drawing.Point(108, 131);
            this.textBoxAddTypeWork.Name = "textBoxAddTypeWork";
            this.textBoxAddTypeWork.Size = new System.Drawing.Size(100, 20);
            this.textBoxAddTypeWork.TabIndex = 4;
            // 
            // FormEditWork
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(226, 195);
            this.Controls.Add(this.textBoxAddTypeWork);
            this.Controls.Add(this.buttonDeleteWork);
            this.Controls.Add(this.buttonAddTypeWork);
            this.Controls.Add(this.checkedListBoxTypeWork);
            this.Name = "FormEditWork";
            this.Text = "Редактирование типа работы";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBoxTypeWork;
        private System.Windows.Forms.Button buttonAddTypeWork;
        private System.Windows.Forms.Button buttonDeleteWork;
        private System.Windows.Forms.TextBox textBoxAddTypeWork;
    }
}