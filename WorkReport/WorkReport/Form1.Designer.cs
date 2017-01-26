namespace WorkReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtfile = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnrunreport = new System.Windows.Forms.Button();
            this.frmdate = new System.Windows.Forms.DateTimePicker();
            this.todate = new System.Windows.Forms.DateTimePicker();
            this.progrssbar = new System.Windows.Forms.ProgressBar();
            this.myBGWorker = new System.ComponentModel.BackgroundWorker();
            this.lblprgrss = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblmsg = new System.Windows.Forms.LinkLabel();
            this.label3 = new System.Windows.Forms.Label();
            this.isdisplayweek = new System.Windows.Forms.CheckBox();
            this.txtmsg = new System.Windows.Forms.TextBox();
            this.genratereport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(98, 95);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "From Date";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(98, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "To Date";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtfile
            // 
            this.txtfile.Location = new System.Drawing.Point(247, 185);
            this.txtfile.Name = "txtfile";
            this.txtfile.Size = new System.Drawing.Size(145, 20);
            this.txtfile.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(101, 185);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "Choose File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnrunreport
            // 
            this.btnrunreport.Location = new System.Drawing.Point(247, 262);
            this.btnrunreport.Name = "btnrunreport";
            this.btnrunreport.Size = new System.Drawing.Size(87, 31);
            this.btnrunreport.TabIndex = 7;
            this.btnrunreport.Text = "Dump Report";
            this.btnrunreport.UseVisualStyleBackColor = true;
            this.btnrunreport.Click += new System.EventHandler(this.btnrunreport_Click);
            // 
            // frmdate
            // 
            this.frmdate.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.frmdate.CustomFormat = "dd/MM/yyyy";
            this.frmdate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.frmdate.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.frmdate.Location = new System.Drawing.Point(247, 88);
            this.frmdate.Name = "frmdate";
            this.frmdate.Size = new System.Drawing.Size(145, 20);
            this.frmdate.TabIndex = 8;
            this.frmdate.Value = new System.DateTime(2017, 1, 6, 0, 0, 0, 0);
            // 
            // todate
            // 
            this.todate.CustomFormat = "dd/MM/yyyy";
            this.todate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.todate.Location = new System.Drawing.Point(247, 131);
            this.todate.Name = "todate";
            this.todate.Size = new System.Drawing.Size(145, 20);
            this.todate.TabIndex = 9;
            // 
            // progrssbar
            // 
            this.progrssbar.Location = new System.Drawing.Point(136, 299);
            this.progrssbar.Name = "progrssbar";
            this.progrssbar.Size = new System.Drawing.Size(301, 29);
            this.progrssbar.TabIndex = 10;
            // 
            // myBGWorker
            // 
            this.myBGWorker.WorkerReportsProgress = true;
            this.myBGWorker.WorkerSupportsCancellation = true;
            this.myBGWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.myBGWorker_DoWork);
            this.myBGWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.myBGWorker_ProgressChanged);
            this.myBGWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.myBGWorker_RunWorkerCompleted);
            // 
            // lblprgrss
            // 
            this.lblprgrss.AutoSize = true;
            this.lblprgrss.Location = new System.Drawing.Point(98, 313);
            this.lblprgrss.Name = "lblprgrss";
            this.lblprgrss.Size = new System.Drawing.Size(0, 13);
            this.lblprgrss.TabIndex = 11;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(0, -2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(568, 64);
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // lblmsg
            // 
            this.lblmsg.ActiveLinkColor = System.Drawing.Color.ForestGreen;
            this.lblmsg.AutoSize = true;
            this.lblmsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblmsg.Location = new System.Drawing.Point(97, 408);
            this.lblmsg.Name = "lblmsg";
            this.lblmsg.Size = new System.Drawing.Size(31, 20);
            this.lblmsg.TabIndex = 13;
            this.lblmsg.TabStop = true;
            this.lblmsg.Text = "AS";
            this.lblmsg.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblmsg_LinkClicked);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(65, 236);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(154, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Display Data For Current Week";
            // 
            // isdisplayweek
            // 
            this.isdisplayweek.AutoSize = true;
            this.isdisplayweek.Location = new System.Drawing.Point(247, 235);
            this.isdisplayweek.Name = "isdisplayweek";
            this.isdisplayweek.Size = new System.Drawing.Size(15, 14);
            this.isdisplayweek.TabIndex = 15;
            this.isdisplayweek.UseVisualStyleBackColor = true;
            // 
            // txtmsg
            // 
            this.txtmsg.Location = new System.Drawing.Point(101, 348);
            this.txtmsg.Multiline = true;
            this.txtmsg.Name = "txtmsg";
            this.txtmsg.Size = new System.Drawing.Size(366, 52);
            this.txtmsg.TabIndex = 16;
            // 
            // genratereport
            // 
            this.genratereport.Location = new System.Drawing.Point(338, 262);
            this.genratereport.Name = "genratereport";
            this.genratereport.Size = new System.Drawing.Size(99, 31);
            this.genratereport.TabIndex = 17;
            this.genratereport.Text = "Generate Report";
            this.genratereport.UseVisualStyleBackColor = true;
            this.genratereport.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(566, 437);
            this.Controls.Add(this.genratereport);
            this.Controls.Add(this.txtmsg);
            this.Controls.Add(this.isdisplayweek);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lblmsg);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lblprgrss);
            this.Controls.Add(this.progrssbar);
            this.Controls.Add(this.todate);
            this.Controls.Add(this.frmdate);
            this.Controls.Add(this.btnrunreport);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtfile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtfile;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnrunreport;
        private System.Windows.Forms.DateTimePicker frmdate;
        private System.Windows.Forms.DateTimePicker todate;
        private System.Windows.Forms.ProgressBar progrssbar;
        private System.ComponentModel.BackgroundWorker myBGWorker;
        private System.Windows.Forms.Label lblprgrss;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.LinkLabel lblmsg;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox isdisplayweek;
        private System.Windows.Forms.TextBox txtmsg;
        private System.Windows.Forms.Button genratereport;
    }
}

