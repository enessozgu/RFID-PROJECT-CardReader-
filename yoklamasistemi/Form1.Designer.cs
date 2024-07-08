namespace yoklamasistemi
{
	partial class Form1
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
			this.label1 = new System.Windows.Forms.Label();
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.buttonAddID = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(14, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(221, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Kart Okuyucu Modeli: ACR1281U-C2";
			// 
			// listBox1
			// 
			this.listBox1.FormattingEnabled = true;
			this.listBox1.ItemHeight = 16;
			this.listBox1.Location = new System.Drawing.Point(14, 43);
			this.listBox1.Name = "listBox1";
			this.listBox1.Size = new System.Drawing.Size(411, 292);
			this.listBox1.TabIndex = 1;
			// 
			// textBox1
			// 
			this.textBox1.Location = new System.Drawing.Point(14, 341);
			this.textBox1.Name = "textBox1";
			this.textBox1.Size = new System.Drawing.Size(285, 22);
			this.textBox1.TabIndex = 3;
			// 
			// buttonAddID
			// 
			this.buttonAddID.Location = new System.Drawing.Point(306, 341);
			this.buttonAddID.Name = "buttonAddID";
			this.buttonAddID.Size = new System.Drawing.Size(119, 25);
			this.buttonAddID.TabIndex = 4;
			this.buttonAddID.Text = "ID Ekle";
			this.buttonAddID.UseVisualStyleBackColor = true;
			this.buttonAddID.Click += new System.EventHandler(this.buttonAddID_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(439, 417);
			this.Controls.Add(this.buttonAddID);
			this.Controls.Add(this.textBox1);
			this.Controls.Add(this.listBox1);
			this.Controls.Add(this.label1);
			this.Name = "Form1";
			this.Text = "RFID Okuyucu";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}





		#endregion

		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ListBox listBox1;
		private System.Windows.Forms.Button buttonAddID;
	}
}

