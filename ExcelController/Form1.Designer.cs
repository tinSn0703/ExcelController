namespace ExcelController
{
	partial class Form1
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows フォーム デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.ActionButton = new System.Windows.Forms.Button();
			this.FIleOpenButton = new System.Windows.Forms.Button();
			this.FileNameTextBox = new System.Windows.Forms.TextBox();
			this.ResultTextBox = new System.Windows.Forms.TextBox();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.SuspendLayout();
			// 
			// ActionButton
			// 
			this.ActionButton.Location = new System.Drawing.Point(12, 85);
			this.ActionButton.Name = "ActionButton";
			this.ActionButton.Size = new System.Drawing.Size(75, 23);
			this.ActionButton.TabIndex = 0;
			this.ActionButton.Text = "実行";
			this.ActionButton.UseVisualStyleBackColor = true;
			this.ActionButton.Click += new System.EventHandler(this.Button2_Click);
			// 
			// FIleOpenButton
			// 
			this.FIleOpenButton.Location = new System.Drawing.Point(197, 37);
			this.FIleOpenButton.Name = "FIleOpenButton";
			this.FIleOpenButton.Size = new System.Drawing.Size(75, 23);
			this.FIleOpenButton.TabIndex = 0;
			this.FIleOpenButton.Text = "開く";
			this.FIleOpenButton.UseVisualStyleBackColor = true;
			this.FIleOpenButton.Click += new System.EventHandler(this.Button1_Click);
			// 
			// FileNameTextBox
			// 
			this.FileNameTextBox.Location = new System.Drawing.Point(12, 41);
			this.FileNameTextBox.Name = "FileNameTextBox";
			this.FileNameTextBox.Size = new System.Drawing.Size(179, 19);
			this.FileNameTextBox.TabIndex = 1;
			this.FileNameTextBox.TextChanged += new System.EventHandler(this.TextBox3_TextChanged);
			// 
			// ResultTextBox
			// 
			this.ResultTextBox.Location = new System.Drawing.Point(12, 114);
			this.ResultTextBox.Multiline = true;
			this.ResultTextBox.Name = "ResultTextBox";
			this.ResultTextBox.Size = new System.Drawing.Size(260, 135);
			this.ResultTextBox.TabIndex = 1;
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenFileDialog1_FileOk);
			// 
			// Form1
			// 
			this.ClientSize = new System.Drawing.Size(284, 261);
			this.Controls.Add(this.FileNameTextBox);
			this.Controls.Add(this.ResultTextBox);
			this.Controls.Add(this.ActionButton);
			this.Controls.Add(this.FIleOpenButton);
			this.Name = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button FIleOpenButton;
		private System.Windows.Forms.Button ActionButton;
		private System.Windows.Forms.TextBox FileNameTextBox;
		private System.Windows.Forms.TextBox ResultTextBox;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
	}
}

