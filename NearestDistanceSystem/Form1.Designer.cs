namespace NearestDistanceSystem
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
            this.btnAddPopulationExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPopulationExcel = new System.Windows.Forms.TextBox();
            this.dgvPopulation = new System.Windows.Forms.DataGridView();
            this.dgvDistance = new System.Windows.Forms.DataGridView();
            this.dgvCapacity = new System.Windows.Forms.DataGridView();
            this.btnAddCapacityExcel = new System.Windows.Forms.Button();
            this.btnAddDistanceExcel = new System.Windows.Forms.Button();
            this.btnSaveData = new System.Windows.Forms.Button();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.btnReCalculate = new System.Windows.Forms.Button();
            this.txtCapacityExcel = new System.Windows.Forms.TextBox();
            this.txtDistanceExcel = new System.Windows.Forms.TextBox();
            this.txtNoLoop = new System.Windows.Forms.TextBox();
            this.txtReleaseCenter = new System.Windows.Forms.TextBox();
            this.txtLoop = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTravelDistance = new System.Windows.Forms.TextBox();
            this.txtTotalAllocated = new System.Windows.Forms.TextBox();
            this.txtTotalUnallocated = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPopulation)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDistance)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCapacity)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAddPopulationExcel
            // 
            this.btnAddPopulationExcel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAddPopulationExcel.Location = new System.Drawing.Point(35, 106);
            this.btnAddPopulationExcel.Name = "btnAddPopulationExcel";
            this.btnAddPopulationExcel.Size = new System.Drawing.Size(158, 65);
            this.btnAddPopulationExcel.TabIndex = 0;
            this.btnAddPopulationExcel.Text = "PopulationExcel";
            this.btnAddPopulationExcel.UseVisualStyleBackColor = true;
            this.btnAddPopulationExcel.Click += new System.EventHandler(this.BtnAddPopulationExcel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 661);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(162, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Release Center Involved";
            // 
            // txtPopulationExcel
            // 
            this.txtPopulationExcel.Location = new System.Drawing.Point(221, 106);
            this.txtPopulationExcel.Multiline = true;
            this.txtPopulationExcel.Name = "txtPopulationExcel";
            this.txtPopulationExcel.Size = new System.Drawing.Size(159, 61);
            this.txtPopulationExcel.TabIndex = 2;
            // 
            // dgvPopulation
            // 
            this.dgvPopulation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvPopulation.Location = new System.Drawing.Point(1279, 719);
            this.dgvPopulation.Name = "dgvPopulation";
            this.dgvPopulation.RowHeadersWidth = 51;
            this.dgvPopulation.RowTemplate.Height = 24;
            this.dgvPopulation.Size = new System.Drawing.Size(240, 150);
            this.dgvPopulation.TabIndex = 3;
            // 
            // dgvDistance
            // 
            this.dgvDistance.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDistance.Location = new System.Drawing.Point(971, 21);
            this.dgvDistance.Name = "dgvDistance";
            this.dgvDistance.RowHeadersWidth = 51;
            this.dgvDistance.RowTemplate.Height = 24;
            this.dgvDistance.Size = new System.Drawing.Size(548, 692);
            this.dgvDistance.TabIndex = 4;
            this.dgvDistance.SelectionChanged += new System.EventHandler(this.dgvDistance_SelectionChanged);
            // 
            // dgvCapacity
            // 
            this.dgvCapacity.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvCapacity.Location = new System.Drawing.Point(971, 719);
            this.dgvCapacity.Name = "dgvCapacity";
            this.dgvCapacity.RowHeadersWidth = 51;
            this.dgvCapacity.RowTemplate.Height = 24;
            this.dgvCapacity.Size = new System.Drawing.Size(240, 150);
            this.dgvCapacity.TabIndex = 5;
            // 
            // btnAddCapacityExcel
            // 
            this.btnAddCapacityExcel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAddCapacityExcel.Location = new System.Drawing.Point(35, 196);
            this.btnAddCapacityExcel.Name = "btnAddCapacityExcel";
            this.btnAddCapacityExcel.Size = new System.Drawing.Size(158, 65);
            this.btnAddCapacityExcel.TabIndex = 6;
            this.btnAddCapacityExcel.Text = "CapacityExcel";
            this.btnAddCapacityExcel.UseVisualStyleBackColor = true;
            this.btnAddCapacityExcel.Click += new System.EventHandler(this.BtnAddCapacityExcel_Click);
            // 
            // btnAddDistanceExcel
            // 
            this.btnAddDistanceExcel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAddDistanceExcel.Location = new System.Drawing.Point(35, 283);
            this.btnAddDistanceExcel.Name = "btnAddDistanceExcel";
            this.btnAddDistanceExcel.Size = new System.Drawing.Size(158, 65);
            this.btnAddDistanceExcel.TabIndex = 7;
            this.btnAddDistanceExcel.Text = "DistanceExcel";
            this.btnAddDistanceExcel.UseVisualStyleBackColor = true;
            this.btnAddDistanceExcel.Click += new System.EventHandler(this.BtnAddDistanceExcel_Click);
            // 
            // btnSaveData
            // 
            this.btnSaveData.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSaveData.Location = new System.Drawing.Point(128, 547);
            this.btnSaveData.Name = "btnSaveData";
            this.btnSaveData.Size = new System.Drawing.Size(158, 65);
            this.btnSaveData.TabIndex = 8;
            this.btnSaveData.Text = "SaveData";
            this.btnSaveData.UseVisualStyleBackColor = true;
            this.btnSaveData.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCalculate
            // 
            this.btnCalculate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCalculate.Location = new System.Drawing.Point(128, 368);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(158, 65);
            this.btnCalculate.TabIndex = 9;
            this.btnCalculate.Text = "Click Here to Calculate";
            this.btnCalculate.UseVisualStyleBackColor = true;
            // 
            // btnReCalculate
            // 
            this.btnReCalculate.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnReCalculate.Location = new System.Drawing.Point(128, 460);
            this.btnReCalculate.Name = "btnReCalculate";
            this.btnReCalculate.Size = new System.Drawing.Size(158, 65);
            this.btnReCalculate.TabIndex = 10;
            this.btnReCalculate.Text = "Click Here to Re-Calculate";
            this.btnReCalculate.UseVisualStyleBackColor = true;
            this.btnReCalculate.Click += new System.EventHandler(this.btnReCalculate_Click);
            // 
            // txtCapacityExcel
            // 
            this.txtCapacityExcel.Location = new System.Drawing.Point(221, 196);
            this.txtCapacityExcel.Multiline = true;
            this.txtCapacityExcel.Name = "txtCapacityExcel";
            this.txtCapacityExcel.Size = new System.Drawing.Size(159, 61);
            this.txtCapacityExcel.TabIndex = 11;
            // 
            // txtDistanceExcel
            // 
            this.txtDistanceExcel.Location = new System.Drawing.Point(221, 287);
            this.txtDistanceExcel.Multiline = true;
            this.txtDistanceExcel.Name = "txtDistanceExcel";
            this.txtDistanceExcel.Size = new System.Drawing.Size(159, 61);
            this.txtDistanceExcel.TabIndex = 12;
            // 
            // txtNoLoop
            // 
            this.txtNoLoop.Location = new System.Drawing.Point(302, 464);
            this.txtNoLoop.Multiline = true;
            this.txtNoLoop.Name = "txtNoLoop";
            this.txtNoLoop.Size = new System.Drawing.Size(108, 48);
            this.txtNoLoop.TabIndex = 13;
            this.txtNoLoop.DoubleClick += new System.EventHandler(this.txtNoLoop_DoubleClick);
            this.txtNoLoop.Enter += new System.EventHandler(this.txtNoLoop_Enter);
            this.txtNoLoop.Leave += new System.EventHandler(this.txtNoLoop_Leave);
            // 
            // txtReleaseCenter
            // 
            this.txtReleaseCenter.Location = new System.Drawing.Point(30, 700);
            this.txtReleaseCenter.Multiline = true;
            this.txtReleaseCenter.Name = "txtReleaseCenter";
            this.txtReleaseCenter.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtReleaseCenter.Size = new System.Drawing.Size(380, 181);
            this.txtReleaseCenter.TabIndex = 14;
            // 
            // txtLoop
            // 
            this.txtLoop.Location = new System.Drawing.Point(429, 21);
            this.txtLoop.Multiline = true;
            this.txtLoop.Name = "txtLoop";
            this.txtLoop.Size = new System.Drawing.Size(536, 675);
            this.txtLoop.TabIndex = 15;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(425, 718);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 17);
            this.label2.TabIndex = 16;
            this.label2.Text = "Average Travel Distance";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(425, 778);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(102, 17);
            this.label3.TabIndex = 17;
            this.label3.Text = "Total Allocated";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(425, 834);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(119, 17);
            this.label4.TabIndex = 18;
            this.label4.Text = "Total Unallocated";
            // 
            // txtTravelDistance
            // 
            this.txtTravelDistance.Location = new System.Drawing.Point(622, 702);
            this.txtTravelDistance.Multiline = true;
            this.txtTravelDistance.Name = "txtTravelDistance";
            this.txtTravelDistance.Size = new System.Drawing.Size(226, 48);
            this.txtTravelDistance.TabIndex = 19;
            // 
            // txtTotalAllocated
            // 
            this.txtTotalAllocated.Location = new System.Drawing.Point(622, 759);
            this.txtTotalAllocated.Multiline = true;
            this.txtTotalAllocated.Name = "txtTotalAllocated";
            this.txtTotalAllocated.Size = new System.Drawing.Size(226, 48);
            this.txtTotalAllocated.TabIndex = 20;
            // 
            // txtTotalUnallocated
            // 
            this.txtTotalUnallocated.Location = new System.Drawing.Point(622, 817);
            this.txtTotalUnallocated.Multiline = true;
            this.txtTotalUnallocated.Name = "txtTotalUnallocated";
            this.txtTotalUnallocated.Size = new System.Drawing.Size(226, 48);
            this.txtTotalUnallocated.TabIndex = 21;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1572, 897);
            this.Controls.Add(this.txtTotalUnallocated);
            this.Controls.Add(this.txtTotalAllocated);
            this.Controls.Add(this.txtTravelDistance);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtLoop);
            this.Controls.Add(this.txtReleaseCenter);
            this.Controls.Add(this.txtNoLoop);
            this.Controls.Add(this.txtDistanceExcel);
            this.Controls.Add(this.txtCapacityExcel);
            this.Controls.Add(this.btnReCalculate);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.btnSaveData);
            this.Controls.Add(this.btnAddDistanceExcel);
            this.Controls.Add(this.btnAddCapacityExcel);
            this.Controls.Add(this.dgvCapacity);
            this.Controls.Add(this.dgvDistance);
            this.Controls.Add(this.dgvPopulation);
            this.Controls.Add(this.txtPopulationExcel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAddPopulationExcel);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Main Page";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPopulation)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDistance)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCapacity)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAddPopulationExcel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPopulationExcel;
        private System.Windows.Forms.DataGridView dgvPopulation;
        private System.Windows.Forms.DataGridView dgvDistance;
        private System.Windows.Forms.DataGridView dgvCapacity;
        private System.Windows.Forms.Button btnAddCapacityExcel;
        private System.Windows.Forms.Button btnAddDistanceExcel;
        private System.Windows.Forms.Button btnSaveData;
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.Button btnReCalculate;
        private System.Windows.Forms.TextBox txtCapacityExcel;
        private System.Windows.Forms.TextBox txtDistanceExcel;
        private System.Windows.Forms.TextBox txtNoLoop;
        private System.Windows.Forms.TextBox txtReleaseCenter;
        private System.Windows.Forms.TextBox txtLoop;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtTravelDistance;
        private System.Windows.Forms.TextBox txtTotalAllocated;
        private System.Windows.Forms.TextBox txtTotalUnallocated;
    }
}

