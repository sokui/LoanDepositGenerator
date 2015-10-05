namespace LoanDepositGenerator
{
    partial class generatorForm
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
            this.components = new System.ComponentModel.Container();
            this.inputBtn = new System.Windows.Forms.Button();
            this.inputPathTxtBox = new System.Windows.Forms.TextBox();
            this.startBtn = new System.Windows.Forms.Button();
            this.processRBox = new System.Windows.Forms.RichTextBox();
            this.inputToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.stopBtn = new System.Windows.Forms.Button();
            this.outputToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.startToolTip = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // inputBtn
            // 
            this.inputBtn.Location = new System.Drawing.Point(12, 10);
            this.inputBtn.Name = "inputBtn";
            this.inputBtn.Size = new System.Drawing.Size(83, 24);
            this.inputBtn.TabIndex = 0;
            this.inputBtn.Text = "Daily Folder...";
            this.inputToolTip.SetToolTip(this.inputBtn, "Select an input folder");
            this.inputBtn.UseVisualStyleBackColor = true;
            this.inputBtn.Click += new System.EventHandler(this.inputBtn_Click);
            // 
            // inputPathTxtBox
            // 
            this.inputPathTxtBox.Location = new System.Drawing.Point(112, 12);
            this.inputPathTxtBox.Name = "inputPathTxtBox";
            this.inputPathTxtBox.Size = new System.Drawing.Size(873, 20);
            this.inputPathTxtBox.TabIndex = 2;
            // 
            // startBtn
            // 
            this.startBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.startBtn.Location = new System.Drawing.Point(12, 40);
            this.startBtn.Name = "startBtn";
            this.startBtn.Size = new System.Drawing.Size(75, 26);
            this.startBtn.TabIndex = 4;
            this.startBtn.Text = "Start";
            this.inputToolTip.SetToolTip(this.startBtn, "Begin generating");
            this.startBtn.UseVisualStyleBackColor = true;
            this.startBtn.Click += new System.EventHandler(this.startBtn_Click);
            // 
            // processRBox
            // 
            this.processRBox.Location = new System.Drawing.Point(12, 70);
            this.processRBox.Name = "processRBox";
            this.processRBox.Size = new System.Drawing.Size(973, 236);
            this.processRBox.TabIndex = 5;
            this.processRBox.Text = "";
            // 
            // stopBtn
            // 
            this.stopBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.stopBtn.Location = new System.Drawing.Point(93, 40);
            this.stopBtn.Name = "stopBtn";
            this.stopBtn.Size = new System.Drawing.Size(75, 26);
            this.stopBtn.TabIndex = 6;
            this.stopBtn.Text = "Stop";
            this.inputToolTip.SetToolTip(this.stopBtn, "Begin generating");
            this.stopBtn.UseVisualStyleBackColor = true;
            this.stopBtn.Click += new System.EventHandler(this.stopBtn_Click);
            // 
            // generatorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(997, 320);
            this.Controls.Add(this.stopBtn);
            this.Controls.Add(this.processRBox);
            this.Controls.Add(this.startBtn);
            this.Controls.Add(this.inputPathTxtBox);
            this.Controls.Add(this.inputBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "generatorForm";
            this.Text = "Deposit and loan Generator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button inputBtn;
        private System.Windows.Forms.TextBox inputPathTxtBox;
        private System.Windows.Forms.Button startBtn;
        private System.Windows.Forms.RichTextBox processRBox;
        private System.Windows.Forms.ToolTip inputToolTip;
        private System.Windows.Forms.ToolTip outputToolTip;
        private System.Windows.Forms.ToolTip startToolTip;
        private System.Windows.Forms.Button stopBtn;
    }
}

