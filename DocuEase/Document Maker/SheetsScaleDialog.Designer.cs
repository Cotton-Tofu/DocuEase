namespace Document_Maker
{
    partial class SheetsScaleDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SheetsScaleDialog));
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonCheckBox1 = new ComponentFactory.Krypton.Toolkit.KryptonCheckBox();
            this.kryptonPalette1 = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.kryptonLabel42 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonButton15 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            this.kryptonTrackBar1 = new ComponentFactory.Krypton.Toolkit.KryptonTrackBar();
            this.kryptonButton14 = new ComponentFactory.Krypton.Toolkit.KryptonButton();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.kryptonCheckBox1);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel42);
            this.kryptonPanel1.Controls.Add(this.kryptonButton15);
            this.kryptonPanel1.Controls.Add(this.kryptonTrackBar1);
            this.kryptonPanel1.Controls.Add(this.kryptonButton14);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Palette = this.kryptonPalette1;
            this.kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel1.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.PanelRibbonInactive;
            this.kryptonPanel1.Size = new System.Drawing.Size(537, 51);
            this.kryptonPanel1.TabIndex = 3;
            // 
            // kryptonCheckBox1
            // 
            this.kryptonCheckBox1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonCheckBox1.Location = new System.Drawing.Point(425, 14);
            this.kryptonCheckBox1.Name = "kryptonCheckBox1";
            this.kryptonCheckBox1.Palette = this.kryptonPalette1;
            this.kryptonCheckBox1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonCheckBox1.Size = new System.Drawing.Size(98, 20);
            this.kryptonCheckBox1.TabIndex = 17;
            this.kryptonCheckBox1.Values.Text = "最前面に表示";
            this.kryptonCheckBox1.CheckedChanged += new System.EventHandler(this.kryptonCheckBox1_CheckedChanged);
            // 
            // kryptonLabel42
            // 
            this.kryptonLabel42.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalPanel;
            this.kryptonLabel42.Location = new System.Drawing.Point(13, 14);
            this.kryptonLabel42.Name = "kryptonLabel42";
            this.kryptonLabel42.Palette = this.kryptonPalette1;
            this.kryptonLabel42.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel42.Size = new System.Drawing.Size(34, 20);
            this.kryptonLabel42.TabIndex = 16;
            this.kryptonLabel42.Values.Text = "10%";
            // 
            // kryptonButton15
            // 
            this.kryptonButton15.Enabled = false;
            this.kryptonButton15.Location = new System.Drawing.Point(54, 11);
            this.kryptonButton15.Name = "kryptonButton15";
            this.kryptonButton15.Palette = this.kryptonPalette1;
            this.kryptonButton15.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton15.Size = new System.Drawing.Size(25, 25);
            this.kryptonButton15.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonButton15.StateCommon.Border.Rounding = 11;
            this.kryptonButton15.TabIndex = 15;
            this.kryptonButton15.Values.Image = ((System.Drawing.Image)(resources.GetObject("kryptonButton15.Values.Image")));
            this.kryptonButton15.Values.Text = "-";
            this.kryptonButton15.Click += new System.EventHandler(this.kryptonButton15_Click);
            // 
            // kryptonTrackBar1
            // 
            this.kryptonTrackBar1.BackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.PanelRibbonInactive;
            this.kryptonTrackBar1.DrawBackground = true;
            this.kryptonTrackBar1.Location = new System.Drawing.Point(85, 11);
            this.kryptonTrackBar1.Name = "kryptonTrackBar1";
            this.kryptonTrackBar1.Palette = this.kryptonPalette1;
            this.kryptonTrackBar1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonTrackBar1.Size = new System.Drawing.Size(302, 27);
            this.kryptonTrackBar1.TabIndex = 14;
            this.kryptonTrackBar1.TickStyle = System.Windows.Forms.TickStyle.TopLeft;
            this.kryptonTrackBar1.ValueChanged += new System.EventHandler(this.kryptonTrackBar1_ValueChanged);
            // 
            // kryptonButton14
            // 
            this.kryptonButton14.Location = new System.Drawing.Point(393, 12);
            this.kryptonButton14.Name = "kryptonButton14";
            this.kryptonButton14.Palette = this.kryptonPalette1;
            this.kryptonButton14.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonButton14.Size = new System.Drawing.Size(25, 25);
            this.kryptonButton14.StateCommon.Border.DrawBorders = ((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders)((((ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Top | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Left) 
            | ComponentFactory.Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonButton14.StateCommon.Border.Rounding = 11;
            this.kryptonButton14.TabIndex = 13;
            this.kryptonButton14.Values.Image = ((System.Drawing.Image)(resources.GetObject("kryptonButton14.Values.Image")));
            this.kryptonButton14.Values.Text = "+";
            this.kryptonButton14.Click += new System.EventHandler(this.kryptonButton14_Click);
            // 
            // SheetsScaleDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(537, 51);
            this.Controls.Add(this.kryptonPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "SheetsScaleDialog";
            this.Palette = this.kryptonPalette1;
            this.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.Text = "シートの拡大率";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SheetsScaleDialog_FormClosing);
            this.Load += new System.EventHandler(this.SheetsScaleDialog_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.SheetsScaleDialog_Paint);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            this.kryptonPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton15;
        private ComponentFactory.Krypton.Toolkit.KryptonTrackBar kryptonTrackBar1;
        private ComponentFactory.Krypton.Toolkit.KryptonCheckBox kryptonCheckBox1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel42;
        private ComponentFactory.Krypton.Toolkit.KryptonButton kryptonButton14;
        private ComponentFactory.Krypton.Toolkit.KryptonPalette kryptonPalette1;
    }
}