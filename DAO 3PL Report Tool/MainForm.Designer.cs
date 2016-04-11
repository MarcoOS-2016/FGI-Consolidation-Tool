namespace DAO_3PL_Report_Tool
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.dao3PLReport_Tab_startButton = new System.Windows.Forms.Button();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.apj3PLReport_Tab_SnPPartListFileButton = new System.Windows.Forms.Button();
            this.apj3PLReport_Tab_SnPPartListFileTextBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.apj3PLReport_Tab_StartButton = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.apj3PLReport_Tab_OutputFolderButton = new System.Windows.Forms.Button();
            this.apj3PLReport_Tab_SourceReportTextBox = new System.Windows.Forms.TextBox();
            this.apj3PLReport_Tab_OutputFolderTextBox = new System.Windows.Forms.TextBox();
            this.apj3PLReport_Tab_SourceReportButton = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dao3PLReport_Tab_outputFolderButton = new System.Windows.Forms.Button();
            this.dao3PLReport_Tab_sourceFolderTextBox = new System.Windows.Forms.TextBox();
            this.dao3PLReport_Tab_outputFolderTextBox = new System.Windows.Forms.TextBox();
            this.dao3PLReport_Tab_sourceFolderButton = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dao3PLReport_Tab_beginDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.dao3PLReport_Tab_snapShotDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.dao3PLReport_Tab_endDateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.generateAging_Tab_InTransitReportOpenButton = new System.Windows.Forms.Button();
            this.generateAging_Tab_InTransitReportTextBox = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.generateAging_Tab_startButton = new System.Windows.Forms.Button();
            this.generateAging_Tab_OnhandReportOpenButton = new System.Windows.Forms.Button();
            this.generateAging_Tab_OnhandReportTextBox = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.generateAging_Tab_DetailsReportOpenButton = new System.Windows.Forms.Button();
            this.generateAging_Tab_DetailsReportFolderTextBox = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.generateAging_Tab_AgingReportOpenButton = new System.Windows.Forms.Button();
            this.generateAging_Tab_AgingReportFolderTextBox = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.generateAging_Tab_WeekNumberComboBox = new System.Windows.Forms.ComboBox();
            this.generateAging_Tab_OutputFolderOpenButton = new System.Windows.Forms.Button();
            this.generateAging_Tab_OutputFolderTextBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.historicalAgingDataMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dao3PLReport_Tab_startButton
            // 
            this.dao3PLReport_Tab_startButton.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.dao3PLReport_Tab_startButton.Location = new System.Drawing.Point(607, 225);
            this.dao3PLReport_Tab_startButton.Name = "dao3PLReport_Tab_startButton";
            this.dao3PLReport_Tab_startButton.Size = new System.Drawing.Size(100, 30);
            this.dao3PLReport_Tab_startButton.TabIndex = 59;
            this.dao3PLReport_Tab_startButton.Text = "Start";
            this.dao3PLReport_Tab_startButton.UseVisualStyleBackColor = true;
            this.dao3PLReport_Tab_startButton.Click += new System.EventHandler(this.startButton_Click);
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip.Location = new System.Drawing.Point(0, 434);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(752, 22);
            this.statusStrip.TabIndex = 63;
            this.statusStrip.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.tabControl1.Location = new System.Drawing.Point(0, 34);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(752, 397);
            this.tabControl1.TabIndex = 64;
            // 
            // tabPage4
            // 
            this.tabPage4.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tabPage4.Controls.Add(this.groupBox3);
            this.tabPage4.Location = new System.Drawing.Point(4, 25);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(744, 368);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Consolidate APJ 3PL Reports";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_SnPPartListFileButton);
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_SnPPartListFileTextBox);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_StartButton);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_OutputFolderButton);
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_SourceReportTextBox);
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_OutputFolderTextBox);
            this.groupBox3.Controls.Add(this.apj3PLReport_Tab_SourceReportButton);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Location = new System.Drawing.Point(9, 6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(725, 212);
            this.groupBox3.TabIndex = 64;
            this.groupBox3.TabStop = false;
            // 
            // apj3PLReport_Tab_SnPPartListFileButton
            // 
            this.apj3PLReport_Tab_SnPPartListFileButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.apj3PLReport_Tab_SnPPartListFileButton.Image = ((System.Drawing.Image)(resources.GetObject("apj3PLReport_Tab_SnPPartListFileButton.Image")));
            this.apj3PLReport_Tab_SnPPartListFileButton.Location = new System.Drawing.Point(679, 72);
            this.apj3PLReport_Tab_SnPPartListFileButton.Name = "apj3PLReport_Tab_SnPPartListFileButton";
            this.apj3PLReport_Tab_SnPPartListFileButton.Size = new System.Drawing.Size(28, 25);
            this.apj3PLReport_Tab_SnPPartListFileButton.TabIndex = 77;
            this.apj3PLReport_Tab_SnPPartListFileButton.UseVisualStyleBackColor = true;
            this.apj3PLReport_Tab_SnPPartListFileButton.Click += new System.EventHandler(this.apj3PLReport_Tab_SnPPartListFileButton_Click);
            // 
            // apj3PLReport_Tab_SnPPartListFileTextBox
            // 
            this.apj3PLReport_Tab_SnPPartListFileTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.apj3PLReport_Tab_SnPPartListFileTextBox.Location = new System.Drawing.Point(174, 73);
            this.apj3PLReport_Tab_SnPPartListFileTextBox.Name = "apj3PLReport_Tab_SnPPartListFileTextBox";
            this.apj3PLReport_Tab_SnPPartListFileTextBox.Size = new System.Drawing.Size(496, 23);
            this.apj3PLReport_Tab_SnPPartListFileTextBox.TabIndex = 76;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(59, 76);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(109, 16);
            this.label6.TabIndex = 75;
            this.label6.Text = "SnP Part List File:";
            // 
            // apj3PLReport_Tab_StartButton
            // 
            this.apj3PLReport_Tab_StartButton.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.apj3PLReport_Tab_StartButton.Location = new System.Drawing.Point(607, 164);
            this.apj3PLReport_Tab_StartButton.Name = "apj3PLReport_Tab_StartButton";
            this.apj3PLReport_Tab_StartButton.Size = new System.Drawing.Size(100, 30);
            this.apj3PLReport_Tab_StartButton.TabIndex = 59;
            this.apj3PLReport_Tab_StartButton.Text = "Start";
            this.apj3PLReport_Tab_StartButton.UseVisualStyleBackColor = true;
            this.apj3PLReport_Tab_StartButton.Click += new System.EventHandler(this.apj3PLReport_Tab_StartButton_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label7.ForeColor = System.Drawing.Color.Black;
            this.label7.Location = new System.Drawing.Point(11, 33);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(159, 16);
            this.label7.TabIndex = 63;
            this.label7.Text = "APJ Source Report Folder:";
            // 
            // apj3PLReport_Tab_OutputFolderButton
            // 
            this.apj3PLReport_Tab_OutputFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.apj3PLReport_Tab_OutputFolderButton.Image = ((System.Drawing.Image)(resources.GetObject("apj3PLReport_Tab_OutputFolderButton.Image")));
            this.apj3PLReport_Tab_OutputFolderButton.Location = new System.Drawing.Point(679, 116);
            this.apj3PLReport_Tab_OutputFolderButton.Name = "apj3PLReport_Tab_OutputFolderButton";
            this.apj3PLReport_Tab_OutputFolderButton.Size = new System.Drawing.Size(28, 25);
            this.apj3PLReport_Tab_OutputFolderButton.TabIndex = 74;
            this.apj3PLReport_Tab_OutputFolderButton.UseVisualStyleBackColor = true;
            this.apj3PLReport_Tab_OutputFolderButton.Click += new System.EventHandler(this.apj3PLReport_Tab_OutputFolderButton_Click);
            // 
            // apj3PLReport_Tab_SourceReportTextBox
            // 
            this.apj3PLReport_Tab_SourceReportTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.apj3PLReport_Tab_SourceReportTextBox.Location = new System.Drawing.Point(174, 30);
            this.apj3PLReport_Tab_SourceReportTextBox.Name = "apj3PLReport_Tab_SourceReportTextBox";
            this.apj3PLReport_Tab_SourceReportTextBox.Size = new System.Drawing.Size(496, 23);
            this.apj3PLReport_Tab_SourceReportTextBox.TabIndex = 64;
            // 
            // apj3PLReport_Tab_OutputFolderTextBox
            // 
            this.apj3PLReport_Tab_OutputFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.apj3PLReport_Tab_OutputFolderTextBox.Location = new System.Drawing.Point(174, 117);
            this.apj3PLReport_Tab_OutputFolderTextBox.Name = "apj3PLReport_Tab_OutputFolderTextBox";
            this.apj3PLReport_Tab_OutputFolderTextBox.Size = new System.Drawing.Size(496, 23);
            this.apj3PLReport_Tab_OutputFolderTextBox.TabIndex = 73;
            // 
            // apj3PLReport_Tab_SourceReportButton
            // 
            this.apj3PLReport_Tab_SourceReportButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.apj3PLReport_Tab_SourceReportButton.Image = ((System.Drawing.Image)(resources.GetObject("apj3PLReport_Tab_SourceReportButton.Image")));
            this.apj3PLReport_Tab_SourceReportButton.Location = new System.Drawing.Point(679, 29);
            this.apj3PLReport_Tab_SourceReportButton.Name = "apj3PLReport_Tab_SourceReportButton";
            this.apj3PLReport_Tab_SourceReportButton.Size = new System.Drawing.Size(28, 25);
            this.apj3PLReport_Tab_SourceReportButton.TabIndex = 65;
            this.apj3PLReport_Tab_SourceReportButton.UseVisualStyleBackColor = true;
            this.apj3PLReport_Tab_SourceReportButton.Click += new System.EventHandler(this.apj3PLReport_Tab_SourceReportButton_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label8.ForeColor = System.Drawing.Color.Black;
            this.label8.Location = new System.Drawing.Point(79, 120);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(91, 16);
            this.label8.TabIndex = 72;
            this.label8.Text = "Output Folder:";
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(744, 368);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = " Consolidate DAO 3PL Reports";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_startButton);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_outputFolderButton);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_sourceFolderTextBox);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_outputFolderTextBox);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_sourceFolderButton);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_beginDateTimePicker);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_snapShotDateTimePicker);
            this.groupBox1.Controls.Add(this.dao3PLReport_Tab_endDateTimePicker);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Location = new System.Drawing.Point(9, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(725, 274);
            this.groupBox1.TabIndex = 63;
            this.groupBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(411, 30);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(115, 16);
            this.label3.TabIndex = 68;
            this.label3.Text = "End Date of TRAN:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(8, 139);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(164, 16);
            this.label1.TabIndex = 63;
            this.label1.Text = "DAO Source Report Folder:";
            // 
            // dao3PLReport_Tab_outputFolderButton
            // 
            this.dao3PLReport_Tab_outputFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dao3PLReport_Tab_outputFolderButton.Image = ((System.Drawing.Image)(resources.GetObject("dao3PLReport_Tab_outputFolderButton.Image")));
            this.dao3PLReport_Tab_outputFolderButton.Location = new System.Drawing.Point(679, 178);
            this.dao3PLReport_Tab_outputFolderButton.Name = "dao3PLReport_Tab_outputFolderButton";
            this.dao3PLReport_Tab_outputFolderButton.Size = new System.Drawing.Size(28, 25);
            this.dao3PLReport_Tab_outputFolderButton.TabIndex = 74;
            this.dao3PLReport_Tab_outputFolderButton.UseVisualStyleBackColor = true;
            // 
            // dao3PLReport_Tab_sourceFolderTextBox
            // 
            this.dao3PLReport_Tab_sourceFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.dao3PLReport_Tab_sourceFolderTextBox.Location = new System.Drawing.Point(176, 136);
            this.dao3PLReport_Tab_sourceFolderTextBox.Name = "dao3PLReport_Tab_sourceFolderTextBox";
            this.dao3PLReport_Tab_sourceFolderTextBox.Size = new System.Drawing.Size(494, 23);
            this.dao3PLReport_Tab_sourceFolderTextBox.TabIndex = 64;
            // 
            // dao3PLReport_Tab_outputFolderTextBox
            // 
            this.dao3PLReport_Tab_outputFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.dao3PLReport_Tab_outputFolderTextBox.Location = new System.Drawing.Point(176, 179);
            this.dao3PLReport_Tab_outputFolderTextBox.Name = "dao3PLReport_Tab_outputFolderTextBox";
            this.dao3PLReport_Tab_outputFolderTextBox.Size = new System.Drawing.Size(494, 23);
            this.dao3PLReport_Tab_outputFolderTextBox.TabIndex = 73;
            // 
            // dao3PLReport_Tab_sourceFolderButton
            // 
            this.dao3PLReport_Tab_sourceFolderButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dao3PLReport_Tab_sourceFolderButton.Image = ((System.Drawing.Image)(resources.GetObject("dao3PLReport_Tab_sourceFolderButton.Image")));
            this.dao3PLReport_Tab_sourceFolderButton.Location = new System.Drawing.Point(679, 135);
            this.dao3PLReport_Tab_sourceFolderButton.Name = "dao3PLReport_Tab_sourceFolderButton";
            this.dao3PLReport_Tab_sourceFolderButton.Size = new System.Drawing.Size(28, 25);
            this.dao3PLReport_Tab_sourceFolderButton.TabIndex = 65;
            this.dao3PLReport_Tab_sourceFolderButton.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(81, 182);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(91, 16);
            this.label5.TabIndex = 72;
            this.label5.Text = "Output Folder:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(47, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(125, 16);
            this.label2.TabIndex = 66;
            this.label2.Text = "Begin Date of TRAN:";
            // 
            // dao3PLReport_Tab_beginDateTimePicker
            // 
            this.dao3PLReport_Tab_beginDateTimePicker.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.dao3PLReport_Tab_beginDateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dao3PLReport_Tab_beginDateTimePicker.Location = new System.Drawing.Point(176, 28);
            this.dao3PLReport_Tab_beginDateTimePicker.Name = "dao3PLReport_Tab_beginDateTimePicker";
            this.dao3PLReport_Tab_beginDateTimePicker.Size = new System.Drawing.Size(138, 23);
            this.dao3PLReport_Tab_beginDateTimePicker.TabIndex = 67;
            // 
            // dao3PLReport_Tab_snapShotDateTimePicker
            // 
            this.dao3PLReport_Tab_snapShotDateTimePicker.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.dao3PLReport_Tab_snapShotDateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dao3PLReport_Tab_snapShotDateTimePicker.Location = new System.Drawing.Point(176, 69);
            this.dao3PLReport_Tab_snapShotDateTimePicker.Name = "dao3PLReport_Tab_snapShotDateTimePicker";
            this.dao3PLReport_Tab_snapShotDateTimePicker.Size = new System.Drawing.Size(138, 23);
            this.dao3PLReport_Tab_snapShotDateTimePicker.TabIndex = 71;
            // 
            // dao3PLReport_Tab_endDateTimePicker
            // 
            this.dao3PLReport_Tab_endDateTimePicker.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.dao3PLReport_Tab_endDateTimePicker.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dao3PLReport_Tab_endDateTimePicker.Location = new System.Drawing.Point(532, 28);
            this.dao3PLReport_Tab_endDateTimePicker.Name = "dao3PLReport_Tab_endDateTimePicker";
            this.dao3PLReport_Tab_endDateTimePicker.Size = new System.Drawing.Size(138, 23);
            this.dao3PLReport_Tab_endDateTimePicker.TabIndex = 69;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(17, 72);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(155, 16);
            this.label4.TabIndex = 70;
            this.label4.Text = "Inventory SnapShot Date:";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(744, 368);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Generate Aging Report";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.generateAging_Tab_InTransitReportOpenButton);
            this.groupBox2.Controls.Add(this.generateAging_Tab_InTransitReportTextBox);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.generateAging_Tab_startButton);
            this.groupBox2.Controls.Add(this.generateAging_Tab_OnhandReportOpenButton);
            this.groupBox2.Controls.Add(this.generateAging_Tab_OnhandReportTextBox);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.generateAging_Tab_DetailsReportOpenButton);
            this.groupBox2.Controls.Add(this.generateAging_Tab_DetailsReportFolderTextBox);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.generateAging_Tab_AgingReportOpenButton);
            this.groupBox2.Controls.Add(this.generateAging_Tab_AgingReportFolderTextBox);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.generateAging_Tab_WeekNumberComboBox);
            this.groupBox2.Controls.Add(this.generateAging_Tab_OutputFolderOpenButton);
            this.groupBox2.Controls.Add(this.generateAging_Tab_OutputFolderTextBox);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Location = new System.Drawing.Point(9, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(725, 356);
            this.groupBox2.TabIndex = 64;
            this.groupBox2.TabStop = false;
            // 
            // generateAging_Tab_InTransitReportOpenButton
            // 
            this.generateAging_Tab_InTransitReportOpenButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateAging_Tab_InTransitReportOpenButton.Image = ((System.Drawing.Image)(resources.GetObject("generateAging_Tab_InTransitReportOpenButton.Image")));
            this.generateAging_Tab_InTransitReportOpenButton.Location = new System.Drawing.Point(682, 199);
            this.generateAging_Tab_InTransitReportOpenButton.Name = "generateAging_Tab_InTransitReportOpenButton";
            this.generateAging_Tab_InTransitReportOpenButton.Size = new System.Drawing.Size(28, 25);
            this.generateAging_Tab_InTransitReportOpenButton.TabIndex = 95;
            this.generateAging_Tab_InTransitReportOpenButton.UseVisualStyleBackColor = true;
            this.generateAging_Tab_InTransitReportOpenButton.Click += new System.EventHandler(this.generateAging_Tab_InTransitReportButton_Click);
            // 
            // generateAging_Tab_InTransitReportTextBox
            // 
            this.generateAging_Tab_InTransitReportTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_InTransitReportTextBox.Location = new System.Drawing.Point(180, 199);
            this.generateAging_Tab_InTransitReportTextBox.Name = "generateAging_Tab_InTransitReportTextBox";
            this.generateAging_Tab_InTransitReportTextBox.Size = new System.Drawing.Size(496, 23);
            this.generateAging_Tab_InTransitReportTextBox.TabIndex = 94;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label14.Location = new System.Drawing.Point(8, 202);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(166, 16);
            this.label14.TabIndex = 93;
            this.label14.Text = "Folder of In-Transit Report:";
            // 
            // generateAging_Tab_startButton
            // 
            this.generateAging_Tab_startButton.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_startButton.Location = new System.Drawing.Point(610, 310);
            this.generateAging_Tab_startButton.Name = "generateAging_Tab_startButton";
            this.generateAging_Tab_startButton.Size = new System.Drawing.Size(100, 30);
            this.generateAging_Tab_startButton.TabIndex = 65;
            this.generateAging_Tab_startButton.Text = "Start";
            this.generateAging_Tab_startButton.UseVisualStyleBackColor = true;
            this.generateAging_Tab_startButton.Click += new System.EventHandler(this.generateAging_Tab_startButton_Click);
            // 
            // generateAging_Tab_OnhandReportOpenButton
            // 
            this.generateAging_Tab_OnhandReportOpenButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateAging_Tab_OnhandReportOpenButton.Image = ((System.Drawing.Image)(resources.GetObject("generateAging_Tab_OnhandReportOpenButton.Image")));
            this.generateAging_Tab_OnhandReportOpenButton.Location = new System.Drawing.Point(682, 157);
            this.generateAging_Tab_OnhandReportOpenButton.Name = "generateAging_Tab_OnhandReportOpenButton";
            this.generateAging_Tab_OnhandReportOpenButton.Size = new System.Drawing.Size(28, 25);
            this.generateAging_Tab_OnhandReportOpenButton.TabIndex = 92;
            this.generateAging_Tab_OnhandReportOpenButton.UseVisualStyleBackColor = true;
            this.generateAging_Tab_OnhandReportOpenButton.Click += new System.EventHandler(this.generateAging_Tab_OnhandReportOpenButton_Click);
            // 
            // generateAging_Tab_OnhandReportTextBox
            // 
            this.generateAging_Tab_OnhandReportTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_OnhandReportTextBox.Location = new System.Drawing.Point(180, 157);
            this.generateAging_Tab_OnhandReportTextBox.Name = "generateAging_Tab_OnhandReportTextBox";
            this.generateAging_Tab_OnhandReportTextBox.Size = new System.Drawing.Size(496, 23);
            this.generateAging_Tab_OnhandReportTextBox.TabIndex = 91;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label13.Location = new System.Drawing.Point(15, 160);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(159, 16);
            this.label13.TabIndex = 90;
            this.label13.Text = "Folder of On-hand Report:";
            // 
            // generateAging_Tab_DetailsReportOpenButton
            // 
            this.generateAging_Tab_DetailsReportOpenButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateAging_Tab_DetailsReportOpenButton.Image = ((System.Drawing.Image)(resources.GetObject("generateAging_Tab_DetailsReportOpenButton.Image")));
            this.generateAging_Tab_DetailsReportOpenButton.Location = new System.Drawing.Point(682, 114);
            this.generateAging_Tab_DetailsReportOpenButton.Name = "generateAging_Tab_DetailsReportOpenButton";
            this.generateAging_Tab_DetailsReportOpenButton.Size = new System.Drawing.Size(28, 25);
            this.generateAging_Tab_DetailsReportOpenButton.TabIndex = 89;
            this.generateAging_Tab_DetailsReportOpenButton.UseVisualStyleBackColor = true;
            this.generateAging_Tab_DetailsReportOpenButton.Click += new System.EventHandler(this.generateAging_Tab_DetailsReportOpenButton_Click);
            // 
            // generateAging_Tab_DetailsReportFolderTextBox
            // 
            this.generateAging_Tab_DetailsReportFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_DetailsReportFolderTextBox.Location = new System.Drawing.Point(180, 114);
            this.generateAging_Tab_DetailsReportFolderTextBox.Name = "generateAging_Tab_DetailsReportFolderTextBox";
            this.generateAging_Tab_DetailsReportFolderTextBox.Size = new System.Drawing.Size(496, 23);
            this.generateAging_Tab_DetailsReportFolderTextBox.TabIndex = 88;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label12.Location = new System.Drawing.Point(26, 117);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(148, 16);
            this.label12.TabIndex = 87;
            this.label12.Text = "Folder of Details Report:";
            // 
            // generateAging_Tab_AgingReportOpenButton
            // 
            this.generateAging_Tab_AgingReportOpenButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateAging_Tab_AgingReportOpenButton.Image = ((System.Drawing.Image)(resources.GetObject("generateAging_Tab_AgingReportOpenButton.Image")));
            this.generateAging_Tab_AgingReportOpenButton.Location = new System.Drawing.Point(682, 72);
            this.generateAging_Tab_AgingReportOpenButton.Name = "generateAging_Tab_AgingReportOpenButton";
            this.generateAging_Tab_AgingReportOpenButton.Size = new System.Drawing.Size(28, 25);
            this.generateAging_Tab_AgingReportOpenButton.TabIndex = 86;
            this.generateAging_Tab_AgingReportOpenButton.UseVisualStyleBackColor = true;
            this.generateAging_Tab_AgingReportOpenButton.Click += new System.EventHandler(this.generateAging_Tab_AgingReportOpenButton_Click);
            // 
            // generateAging_Tab_AgingReportFolderTextBox
            // 
            this.generateAging_Tab_AgingReportFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_AgingReportFolderTextBox.Location = new System.Drawing.Point(180, 72);
            this.generateAging_Tab_AgingReportFolderTextBox.Name = "generateAging_Tab_AgingReportFolderTextBox";
            this.generateAging_Tab_AgingReportFolderTextBox.Size = new System.Drawing.Size(496, 23);
            this.generateAging_Tab_AgingReportFolderTextBox.TabIndex = 85;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label11.Location = new System.Drawing.Point(32, 76);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(142, 16);
            this.label11.TabIndex = 84;
            this.label11.Text = "Folder of Aging Report:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label10.Location = new System.Drawing.Point(42, 33);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(132, 16);
            this.label10.TabIndex = 82;
            this.label10.Text = "Please select Week#:";
            // 
            // generateAging_Tab_WeekNumberComboBox
            // 
            this.generateAging_Tab_WeekNumberComboBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_WeekNumberComboBox.FormattingEnabled = true;
            this.generateAging_Tab_WeekNumberComboBox.Location = new System.Drawing.Point(180, 30);
            this.generateAging_Tab_WeekNumberComboBox.Name = "generateAging_Tab_WeekNumberComboBox";
            this.generateAging_Tab_WeekNumberComboBox.Size = new System.Drawing.Size(156, 24);
            this.generateAging_Tab_WeekNumberComboBox.TabIndex = 81;
            // 
            // generateAging_Tab_OutputFolderOpenButton
            // 
            this.generateAging_Tab_OutputFolderOpenButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generateAging_Tab_OutputFolderOpenButton.Image = ((System.Drawing.Image)(resources.GetObject("generateAging_Tab_OutputFolderOpenButton.Image")));
            this.generateAging_Tab_OutputFolderOpenButton.Location = new System.Drawing.Point(682, 242);
            this.generateAging_Tab_OutputFolderOpenButton.Name = "generateAging_Tab_OutputFolderOpenButton";
            this.generateAging_Tab_OutputFolderOpenButton.Size = new System.Drawing.Size(28, 25);
            this.generateAging_Tab_OutputFolderOpenButton.TabIndex = 80;
            this.generateAging_Tab_OutputFolderOpenButton.UseVisualStyleBackColor = true;
            this.generateAging_Tab_OutputFolderOpenButton.Click += new System.EventHandler(this.generateAging_Tab_OutputFolderOpenButton_Click);
            // 
            // generateAging_Tab_OutputFolderTextBox
            // 
            this.generateAging_Tab_OutputFolderTextBox.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.generateAging_Tab_OutputFolderTextBox.Location = new System.Drawing.Point(180, 242);
            this.generateAging_Tab_OutputFolderTextBox.Name = "generateAging_Tab_OutputFolderTextBox";
            this.generateAging_Tab_OutputFolderTextBox.Size = new System.Drawing.Size(496, 23);
            this.generateAging_Tab_OutputFolderTextBox.TabIndex = 79;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Tahoma", 9.75F);
            this.label9.ForeColor = System.Drawing.Color.Black;
            this.label9.Location = new System.Drawing.Point(83, 245);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(91, 16);
            this.label9.TabIndex = 78;
            this.label9.Text = "Output Folder:";
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripDropDownButton1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(752, 31);
            this.toolStrip1.TabIndex = 65;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.historicalAgingDataMenuItem});
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(76, 28);
            this.toolStripDropDownButton1.Text = "Query";
            // 
            // historicalAgingDataMenuItem
            // 
            this.historicalAgingDataMenuItem.Name = "historicalAgingDataMenuItem";
            this.historicalAgingDataMenuItem.Size = new System.Drawing.Size(186, 22);
            this.historicalAgingDataMenuItem.Text = "Historical Aging Data";
            this.historicalAgingDataMenuItem.Click += new System.EventHandler(this.historicalAgingDataMenuItem_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(752, 456);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.statusStrip);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FGI Consolidation Tool(Ver.04112016)";
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button dao3PLReport_Tab_startButton;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button dao3PLReport_Tab_outputFolderButton;
        private System.Windows.Forms.TextBox dao3PLReport_Tab_sourceFolderTextBox;
        private System.Windows.Forms.TextBox dao3PLReport_Tab_outputFolderTextBox;
        private System.Windows.Forms.Button dao3PLReport_Tab_sourceFolderButton;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dao3PLReport_Tab_beginDateTimePicker;
        private System.Windows.Forms.DateTimePicker dao3PLReport_Tab_snapShotDateTimePicker;
        private System.Windows.Forms.DateTimePicker dao3PLReport_Tab_endDateTimePicker;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button generateAging_Tab_startButton;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button generateAging_Tab_OutputFolderOpenButton;
        private System.Windows.Forms.TextBox generateAging_Tab_OutputFolderTextBox;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripMenuItem historicalAgingDataMenuItem;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox generateAging_Tab_WeekNumberComboBox;
        private System.Windows.Forms.Button generateAging_Tab_DetailsReportOpenButton;
        private System.Windows.Forms.TextBox generateAging_Tab_DetailsReportFolderTextBox;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button generateAging_Tab_AgingReportOpenButton;
        private System.Windows.Forms.TextBox generateAging_Tab_AgingReportFolderTextBox;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button generateAging_Tab_OnhandReportOpenButton;
        private System.Windows.Forms.TextBox generateAging_Tab_OnhandReportTextBox;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button apj3PLReport_Tab_StartButton;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button apj3PLReport_Tab_OutputFolderButton;
        private System.Windows.Forms.TextBox apj3PLReport_Tab_SourceReportTextBox;
        private System.Windows.Forms.TextBox apj3PLReport_Tab_OutputFolderTextBox;
        private System.Windows.Forms.Button apj3PLReport_Tab_SourceReportButton;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button apj3PLReport_Tab_SnPPartListFileButton;
        private System.Windows.Forms.TextBox apj3PLReport_Tab_SnPPartListFileTextBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button generateAging_Tab_InTransitReportOpenButton;
        private System.Windows.Forms.TextBox generateAging_Tab_InTransitReportTextBox;
        private System.Windows.Forms.Label label14;
    }
}

