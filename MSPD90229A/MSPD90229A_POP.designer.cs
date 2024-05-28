namespace CSI.GMES.PD
{
    partial class MSPD90229A_POP
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MSPD90229A_POP));
            this.gridViewEx1 = new JPlatform.Client.Controls6.GridViewEx();
            this.pnTop = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cboMonth = new JPlatform.Client.Controls6.DateEditEx();
            this.labelEx2 = new JPlatform.Client.Controls6.LabelEx();
            this.cboPlant = new JPlatform.Client.Controls6.LookUpEditEx();
            this.lbDateF = new JPlatform.Client.Controls6.LabelEx();
            this.lbLine = new JPlatform.Client.Controls6.LabelEx();
            this.cboLine = new JPlatform.Client.Controls6.LookUpEditEx();
            this.labelEx1 = new JPlatform.Client.Controls6.LabelEx();
            this.cboFactory = new JPlatform.Client.Controls6.LookUpEditEx();
            this.btnReg = new DevExpress.XtraEditors.SimpleButton();
            this.chkAlert = new System.Windows.Forms.CheckBox();
            this.labelEx3 = new JPlatform.Client.Controls6.LabelEx();
            ((System.ComponentModel.ISupportInitialize)(this.FormMessages)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.FoemComboInfo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.BaseTextEditEx.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewEx1)).BeginInit();
            this.pnTop.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.cboMonth.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboMonth.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboPlant.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboLine.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboFactory.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // BaseTextEditEx
            // 
            this.BaseTextEditEx.Properties.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.BaseTextEditEx.Properties.Appearance.Options.UseFont = true;
            // 
            // gridViewEx1
            // 
            this.gridViewEx1.ActionMode = JPlatform.Client.Controls6.ActionMode.View;
            this.gridViewEx1.Name = "gridViewEx1";
            // 
            // pnTop
            // 
            this.pnTop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.pnTop.Controls.Add(this.chkAlert);
            this.pnTop.Controls.Add(this.labelEx3);
            this.pnTop.Controls.Add(this.groupBox1);
            this.pnTop.Controls.Add(this.btnReg);
            this.pnTop.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnTop.Location = new System.Drawing.Point(0, 0);
            this.pnTop.Name = "pnTop";
            this.pnTop.Size = new System.Drawing.Size(571, 116);
            this.pnTop.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cboMonth);
            this.groupBox1.Controls.Add(this.labelEx2);
            this.groupBox1.Controls.Add(this.cboPlant);
            this.groupBox1.Controls.Add(this.lbDateF);
            this.groupBox1.Controls.Add(this.lbLine);
            this.groupBox1.Controls.Add(this.cboLine);
            this.groupBox1.Controls.Add(this.labelEx1);
            this.groupBox1.Controls.Add(this.cboFactory);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(417, 93);
            this.groupBox1.TabIndex = 566;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Information";
            // 
            // cboMonth
            // 
            this.cboMonth.ControlValue = new System.DateTime(2020, 9, 16, 15, 54, 12, 727);
            this.cboMonth.DefaultValue = JPlatform.Client.Controls6.DateEditEx.DefaultValueType.Now;
            this.cboMonth.EditValue = new System.DateTime(2020, 9, 16, 15, 54, 12, 727);
            this.cboMonth.Location = new System.Drawing.Point(72, 23);
            this.cboMonth.Name = "cboMonth";
            this.cboMonth.Properties.AllowBlank = false;
            this.cboMonth.Properties.Appearance.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.cboMonth.Properties.Appearance.Options.UseFont = true;
            this.cboMonth.Properties.BindingField = "FA_DATE";
            this.cboMonth.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.cboMonth.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo),
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.cboMonth.Properties.DisplayFormat.FormatString = "yyyy-MM";
            this.cboMonth.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
            this.cboMonth.Properties.FormatString = "yyyy/MM";
            this.cboMonth.Properties.WordID = "W2019041914081985853";
            this.cboMonth.Properties.WordText = "FA Date";
            this.cboMonth.Size = new System.Drawing.Size(130, 25);
            this.cboMonth.TabIndex = 578;
            this.cboMonth.EditValueChanged += new System.EventHandler(this.cboMonth_EditValueChanged);
            // 
            // labelEx2
            // 
            this.labelEx2.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.labelEx2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.labelEx2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.labelEx2.Location = new System.Drawing.Point(18, 55);
            this.labelEx2.Name = "labelEx2";
            this.labelEx2.Size = new System.Drawing.Size(53, 25);
            this.labelEx2.TabIndex = 572;
            this.labelEx2.Text = "Plant";
            // 
            // cboPlant
            // 
            this.cboPlant.ControlValue = null;
            this.cboPlant.Location = new System.Drawing.Point(72, 55);
            this.cboPlant.Name = "cboPlant";
            this.cboPlant.Properties.AddEmptyRow = false;
            this.cboPlant.Properties.AllowBlank = false;
            this.cboPlant.Properties.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.cboPlant.Properties.Appearance.Options.UseFont = true;
            this.cboPlant.Properties.BindingField = "FA_WC_CD";
            this.cboPlant.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.cboPlant.Properties.NullText = "";
            this.cboPlant.Properties.ShowCodeColumn = false;
            this.cboPlant.Properties.WordID = "W2018061910540582403";
            this.cboPlant.Properties.WordText = "FA_WC_CD";
            this.cboPlant.Size = new System.Drawing.Size(130, 25);
            this.cboPlant.TabIndex = 571;
            this.cboPlant.EditValueChanged += new System.EventHandler(this.cboPlant_EditValueChanged);
            // 
            // lbDateF
            // 
            this.lbDateF.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.lbDateF.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.lbDateF.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.lbDateF.Location = new System.Drawing.Point(18, 23);
            this.lbDateF.Name = "lbDateF";
            this.lbDateF.Size = new System.Drawing.Size(52, 25);
            this.lbDateF.TabIndex = 570;
            this.lbDateF.Text = "Month";
            // 
            // lbLine
            // 
            this.lbLine.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.lbLine.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.lbLine.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.lbLine.Location = new System.Drawing.Point(207, 55);
            this.lbLine.Name = "lbLine";
            this.lbLine.Size = new System.Drawing.Size(65, 25);
            this.lbLine.TabIndex = 563;
            this.lbLine.Text = "Line";
            // 
            // cboLine
            // 
            this.cboLine.ControlValue = null;
            this.cboLine.Location = new System.Drawing.Point(274, 55);
            this.cboLine.Name = "cboLine";
            this.cboLine.Properties.AllowBlank = false;
            this.cboLine.Properties.Appearance.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboLine.Properties.Appearance.Options.UseFont = true;
            this.cboLine.Properties.BeforeEditValue = null;
            this.cboLine.Properties.BindingField = "PLANT_CD";
            this.cboLine.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.cboLine.Properties.NullText = "";
            this.cboLine.Properties.WordText = "PLANT_CD";
            this.cboLine.Size = new System.Drawing.Size(130, 25);
            this.cboLine.TabIndex = 543;
            this.cboLine.EditValueChanged += new System.EventHandler(this.cboLine_EditValueChanged);
            // 
            // labelEx1
            // 
            this.labelEx1.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.labelEx1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.labelEx1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.labelEx1.Location = new System.Drawing.Point(207, 23);
            this.labelEx1.Name = "labelEx1";
            this.labelEx1.Size = new System.Drawing.Size(65, 25);
            this.labelEx1.TabIndex = 561;
            this.labelEx1.Text = "Factory";
            // 
            // cboFactory
            // 
            this.cboFactory.ControlValue = null;
            this.cboFactory.Location = new System.Drawing.Point(274, 23);
            this.cboFactory.Name = "cboFactory";
            this.cboFactory.Properties.AddEmptyRow = false;
            this.cboFactory.Properties.AllowBlank = false;
            this.cboFactory.Properties.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.cboFactory.Properties.Appearance.Options.UseFont = true;
            this.cboFactory.Properties.BindingField = "FA_WC_CD";
            this.cboFactory.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.cboFactory.Properties.NullText = "";
            this.cboFactory.Properties.ShowCodeColumn = false;
            this.cboFactory.Properties.WordID = "W2018061910540582403";
            this.cboFactory.Properties.WordText = "FA_WC_CD";
            this.cboFactory.Size = new System.Drawing.Size(130, 25);
            this.cboFactory.TabIndex = 539;
            this.cboFactory.EditValueChanged += new System.EventHandler(this.cboFactory_EditValueChanged);
            // 
            // btnReg
            // 
            this.btnReg.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.btnReg.Appearance.Options.UseFont = true;
            this.btnReg.Image = ((System.Drawing.Image)(resources.GetObject("btnReg.Image")));
            this.btnReg.Location = new System.Drawing.Point(440, 64);
            this.btnReg.Name = "btnReg";
            this.btnReg.Size = new System.Drawing.Size(120, 28);
            this.btnReg.TabIndex = 198;
            this.btnReg.Text = "Save";
            this.btnReg.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // chkAlert
            // 
            this.chkAlert.AutoSize = true;
            this.chkAlert.Cursor = System.Windows.Forms.Cursors.Hand;
            this.chkAlert.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkAlert.Location = new System.Drawing.Point(442, 41);
            this.chkAlert.Name = "chkAlert";
            this.chkAlert.Size = new System.Drawing.Size(15, 14);
            this.chkAlert.TabIndex = 567;
            this.chkAlert.UseVisualStyleBackColor = true;
            // 
            // labelEx3
            // 
            this.labelEx3.Appearance.Font = new System.Drawing.Font("Calibri", 12F);
            this.labelEx3.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.labelEx3.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.labelEx3.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
            this.labelEx3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.labelEx3.Location = new System.Drawing.Point(451, 34);
            this.labelEx3.Name = "labelEx3";
            this.labelEx3.Padding = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.labelEx3.Size = new System.Drawing.Size(92, 28);
            this.labelEx3.TabIndex = 568;
            this.labelEx3.Text = "Confirm Y/N";
            // 
            // MSPD90229A_POP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(571, 116);
            this.Controls.Add(this.pnTop);
            this.MaximumSize = new System.Drawing.Size(587, 155);
            this.MinimumSize = new System.Drawing.Size(587, 155);
            this.Name = "MSPD90229A_POP";
            this.Text = "Unconfirm Checklist";
            this.Controls.SetChildIndex(this.pnTop, 0);
            ((System.ComponentModel.ISupportInitialize)(this.FormMessages)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.FoemComboInfo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.BaseTextEditEx.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridViewEx1)).EndInit();
            this.pnTop.ResumeLayout(false);
            this.pnTop.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.cboMonth.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboMonth.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboPlant.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboLine.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cboFactory.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private JPlatform.Client.Controls6.GridViewEx gridViewEx1;
        private System.Windows.Forms.Panel pnTop;
        private JPlatform.Client.Controls6.LookUpEditEx cboLine;
        private JPlatform.Client.Controls6.LookUpEditEx cboFactory;
        private JPlatform.Client.Controls6.LabelEx labelEx1;
        private JPlatform.Client.Controls6.LabelEx lbLine;
        private System.Windows.Forms.GroupBox groupBox1;
        private DevExpress.XtraEditors.SimpleButton btnReg;
        private JPlatform.Client.Controls6.LabelEx lbDateF;
        private JPlatform.Client.Controls6.LabelEx labelEx2;
        private JPlatform.Client.Controls6.LookUpEditEx cboPlant;
        private JPlatform.Client.Controls6.DateEditEx cboMonth;
        private System.Windows.Forms.CheckBox chkAlert;
        private JPlatform.Client.Controls6.LabelEx labelEx3;
    }
}

