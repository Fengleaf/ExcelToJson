
namespace ExcelToJson
{
    partial class ExcelToJson
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.OpenFileButton = new System.Windows.Forms.Button();
            this.fileListPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.ExportButton = new System.Windows.Forms.Button();
            this.SingleExportButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Location = new System.Drawing.Point(12, 12);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(149, 50);
            this.OpenFileButton.TabIndex = 0;
            this.OpenFileButton.Text = "開啟檔案";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OnButton1Click);
            // 
            // fileListPanel
            // 
            this.fileListPanel.AutoScroll = true;
            this.fileListPanel.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.fileListPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.fileListPanel.Location = new System.Drawing.Point(167, 12);
            this.fileListPanel.Name = "fileListPanel";
            this.fileListPanel.Padding = new System.Windows.Forms.Padding(4);
            this.fileListPanel.Size = new System.Drawing.Size(275, 162);
            this.fileListPanel.TabIndex = 1;
            this.fileListPanel.WrapContents = false;
            // 
            // ExportButton
            // 
            this.ExportButton.Location = new System.Drawing.Point(12, 124);
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.Size = new System.Drawing.Size(149, 50);
            this.ExportButton.TabIndex = 2;
            this.ExportButton.Text = "輸出全部的Json";
            this.ExportButton.UseVisualStyleBackColor = true;
            this.ExportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // SingleExportButton
            // 
            this.SingleExportButton.Location = new System.Drawing.Point(12, 68);
            this.SingleExportButton.Name = "SingleExportButton";
            this.SingleExportButton.Size = new System.Drawing.Size(149, 50);
            this.SingleExportButton.TabIndex = 3;
            this.SingleExportButton.Text = "輸出選擇的Json";
            this.SingleExportButton.UseVisualStyleBackColor = true;
            this.SingleExportButton.Click += new System.EventHandler(this.SingleExportButton_Click);
            // 
            // ExcelToJson
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(454, 188);
            this.Controls.Add(this.SingleExportButton);
            this.Controls.Add(this.ExportButton);
            this.Controls.Add(this.fileListPanel);
            this.Controls.Add(this.OpenFileButton);
            this.KeyPreview = true;
            this.Name = "ExcelToJson";
            this.Text = "ExcelToJson";
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.Button OpenFileButton;
        private System.Windows.Forms.FlowLayoutPanel fileListPanel;
        private System.Windows.Forms.Button ExportButton;
        private System.Windows.Forms.Button SingleExportButton;
    }
}

