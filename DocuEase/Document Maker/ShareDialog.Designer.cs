namespace Document_Maker
{
    partial class ShareDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ShareDialog));
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonCommandLinkButton3 = new Krypton.Toolkit.KryptonCommandLinkButton();
            this.kryptonCommandLinkButton2 = new Krypton.Toolkit.KryptonCommandLinkButton();
            this.kryptonLabel2 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            this.kryptonPalette1 = new ComponentFactory.Krypton.Toolkit.KryptonPalette(this.components);
            this.kryptonCommandLinkButton1 = new Krypton.Toolkit.KryptonCommandLinkButton();
            this.kryptonLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.kryptonPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Controls.Add(this.kryptonCommandLinkButton3);
            this.kryptonPanel1.Controls.Add(this.kryptonCommandLinkButton2);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel2);
            this.kryptonPanel1.Controls.Add(this.kryptonCommandLinkButton1);
            this.kryptonPanel1.Controls.Add(this.kryptonLabel1);
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.Palette = this.kryptonPalette1;
            this.kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonPanel1.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.PanelRibbonInactive;
            this.kryptonPanel1.Size = new System.Drawing.Size(418, 294);
            this.kryptonPanel1.TabIndex = 7;
            // 
            // kryptonCommandLinkButton3
            // 
            this.kryptonCommandLinkButton3.CommandLinkImageValues.Image = ((System.Drawing.Image)(resources.GetObject("kryptonCommandLinkButton3.CommandLinkImageValues.Image")));
            this.kryptonCommandLinkButton3.CommandLinkTextValues.Description = "Slack と連携し添付ファイルとしてメッセージ送信します。";
            this.kryptonCommandLinkButton3.CommandLinkTextValues.DescriptionTextHAlignment = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton3.CommandLinkTextValues.DescriptionTextVAlignment = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonCommandLinkButton3.CommandLinkTextValues.Heading = "Slack の添付ファイルとして送信";
            this.kryptonCommandLinkButton3.CommandLinkTextValues.HeadingTextHAlignment = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton3.CommandLinkTextValues.HeadingTextVAlignment = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonCommandLinkButton3.Location = new System.Drawing.Point(12, 216);
            this.kryptonCommandLinkButton3.Name = "kryptonCommandLinkButton3";
            this.kryptonCommandLinkButton3.OverrideFocus.Border.Draw = Krypton.Toolkit.InheritBool.True;
            this.kryptonCommandLinkButton3.OverrideFocus.Border.DrawBorders = ((Krypton.Toolkit.PaletteDrawBorders)((((Krypton.Toolkit.PaletteDrawBorders.Top | Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | Krypton.Toolkit.PaletteDrawBorders.Left) 
            | Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonCommandLinkButton3.OverrideFocus.Border.GraphicsHint = Krypton.Toolkit.PaletteGraphicsHint.AntiAlias;
            this.kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            this.kryptonCommandLinkButton3.Size = new System.Drawing.Size(394, 61);
            this.kryptonCommandLinkButton3.StateCommon.Content.LongText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton3.StateCommon.Content.LongText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonCommandLinkButton3.StateCommon.Content.ShortText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton3.StateCommon.Content.ShortText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonCommandLinkButton3.TabIndex = 7;
            this.kryptonCommandLinkButton3.Click += new System.EventHandler(this.kryptonCommandLinkButton3_Click);
            // 
            // kryptonCommandLinkButton2
            // 
            this.kryptonCommandLinkButton2.CommandLinkImageValues.Image = ((System.Drawing.Image)(resources.GetObject("kryptonCommandLinkButton2.CommandLinkImageValues.Image")));
            this.kryptonCommandLinkButton2.CommandLinkTextValues.Description = "Microsoft Teams と連携し添付ファイルとしてメッセージ送信します。";
            this.kryptonCommandLinkButton2.CommandLinkTextValues.DescriptionTextHAlignment = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton2.CommandLinkTextValues.DescriptionTextVAlignment = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonCommandLinkButton2.CommandLinkTextValues.Heading = "Teams の添付ファイルとして送信";
            this.kryptonCommandLinkButton2.CommandLinkTextValues.HeadingTextHAlignment = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton2.CommandLinkTextValues.HeadingTextVAlignment = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonCommandLinkButton2.Location = new System.Drawing.Point(12, 144);
            this.kryptonCommandLinkButton2.Name = "kryptonCommandLinkButton2";
            this.kryptonCommandLinkButton2.OverrideFocus.Border.Draw = Krypton.Toolkit.InheritBool.True;
            this.kryptonCommandLinkButton2.OverrideFocus.Border.DrawBorders = ((Krypton.Toolkit.PaletteDrawBorders)((((Krypton.Toolkit.PaletteDrawBorders.Top | Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | Krypton.Toolkit.PaletteDrawBorders.Left) 
            | Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonCommandLinkButton2.OverrideFocus.Border.GraphicsHint = Krypton.Toolkit.PaletteGraphicsHint.AntiAlias;
            this.kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            this.kryptonCommandLinkButton2.Size = new System.Drawing.Size(394, 61);
            this.kryptonCommandLinkButton2.StateCommon.Content.LongText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton2.StateCommon.Content.LongText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonCommandLinkButton2.StateCommon.Content.ShortText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton2.StateCommon.Content.ShortText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonCommandLinkButton2.TabIndex = 6;
            this.kryptonCommandLinkButton2.Click += new System.EventHandler(this.kryptonCommandLinkButton2_Click);
            // 
            // kryptonLabel2
            // 
            this.kryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.BoldPanel;
            this.kryptonLabel2.Location = new System.Drawing.Point(12, 47);
            this.kryptonLabel2.Name = "kryptonLabel2";
            this.kryptonLabel2.Palette = this.kryptonPalette1;
            this.kryptonLabel2.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel2.Size = new System.Drawing.Size(58, 20);
            this.kryptonLabel2.TabIndex = 5;
            this.kryptonLabel2.Values.Text = "(表題名)";
            // 
            // kryptonCommandLinkButton1
            // 
            this.kryptonCommandLinkButton1.CommandLinkImageValues.Image = ((System.Drawing.Image)(resources.GetObject("kryptonCommandLinkButton1.CommandLinkImageValues.Image")));
            this.kryptonCommandLinkButton1.CommandLinkTextValues.Description = "Outlookのメール機能と連携し添付ファイルとしてメール送信します。";
            this.kryptonCommandLinkButton1.CommandLinkTextValues.DescriptionTextHAlignment = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton1.CommandLinkTextValues.DescriptionTextVAlignment = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonCommandLinkButton1.CommandLinkTextValues.Heading = "Outlook の添付ファイルとして送信";
            this.kryptonCommandLinkButton1.CommandLinkTextValues.HeadingTextHAlignment = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton1.CommandLinkTextValues.HeadingTextVAlignment = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonCommandLinkButton1.Location = new System.Drawing.Point(12, 73);
            this.kryptonCommandLinkButton1.Name = "kryptonCommandLinkButton1";
            this.kryptonCommandLinkButton1.OverrideFocus.Border.Draw = Krypton.Toolkit.InheritBool.True;
            this.kryptonCommandLinkButton1.OverrideFocus.Border.DrawBorders = ((Krypton.Toolkit.PaletteDrawBorders)((((Krypton.Toolkit.PaletteDrawBorders.Top | Krypton.Toolkit.PaletteDrawBorders.Bottom) 
            | Krypton.Toolkit.PaletteDrawBorders.Left) 
            | Krypton.Toolkit.PaletteDrawBorders.Right)));
            this.kryptonCommandLinkButton1.OverrideFocus.Border.GraphicsHint = Krypton.Toolkit.PaletteGraphicsHint.AntiAlias;
            this.kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            this.kryptonCommandLinkButton1.Size = new System.Drawing.Size(394, 61);
            this.kryptonCommandLinkButton1.StateCommon.Content.LongText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton1.StateCommon.Content.LongText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonCommandLinkButton1.StateCommon.Content.ShortText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Near;
            this.kryptonCommandLinkButton1.StateCommon.Content.ShortText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonCommandLinkButton1.TabIndex = 4;
            this.kryptonCommandLinkButton1.Click += new System.EventHandler(this.kryptonCommandLinkButton1_Click);
            // 
            // kryptonLabel1
            // 
            this.kryptonLabel1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.TitlePanel;
            this.kryptonLabel1.Location = new System.Drawing.Point(12, 12);
            this.kryptonLabel1.Name = "kryptonLabel1";
            this.kryptonLabel1.Palette = this.kryptonPalette1;
            this.kryptonLabel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.kryptonLabel1.Size = new System.Drawing.Size(147, 29);
            this.kryptonLabel1.TabIndex = 3;
            this.kryptonLabel1.Values.Text = "共有方法の選択";
            // 
            // ShareDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(418, 294);
            this.Controls.Add(this.kryptonPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ShareDialog";
            this.Palette = this.kryptonPalette1;
            this.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Custom;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "文書の共有";
            this.Load += new System.EventHandler(this.ShareDialog_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.kryptonPanel1.ResumeLayout(false);
            this.kryptonPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel2;
        private Krypton.Toolkit.KryptonCommandLinkButton kryptonCommandLinkButton1;
        private Krypton.Toolkit.KryptonCommandLinkButton kryptonCommandLinkButton2;
        private Krypton.Toolkit.KryptonCommandLinkButton kryptonCommandLinkButton3;
        private ComponentFactory.Krypton.Toolkit.KryptonPalette kryptonPalette1;
    }
}