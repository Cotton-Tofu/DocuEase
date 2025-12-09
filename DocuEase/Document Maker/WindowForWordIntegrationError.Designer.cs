namespace Document_Maker
{
    partial class WindowForWordIntegrationError
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
            this.wizardPageContainer1 = new AeroWizard.WizardPageContainer();
            this.wizardControl1 = new AeroWizard.WizardControl();
            this.wizardPage1 = new AeroWizard.WizardPage();
            this.kryptonPanel1 = new ComponentFactory.Krypton.Toolkit.KryptonPanel();
            this.kryptonLabel1 = new ComponentFactory.Krypton.Toolkit.KryptonLabel();
            ((System.ComponentModel.ISupportInitialize)(this.wizardPageContainer1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wizardControl1)).BeginInit();
            this.wizardPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).BeginInit();
            this.SuspendLayout();
            // 
            // wizardPageContainer1
            // 
            this.wizardPageContainer1.BackButton = null;
            this.wizardPageContainer1.CancelButton = null;
            this.wizardPageContainer1.Location = new System.Drawing.Point(100, 65);
            this.wizardPageContainer1.Name = "wizardPageContainer1";
            this.wizardPageContainer1.NextButton = null;
            this.wizardPageContainer1.Size = new System.Drawing.Size(75, 115);
            this.wizardPageContainer1.TabIndex = 0;
            // 
            // wizardControl1
            // 
            this.wizardControl1.ClassicStyle = AeroWizard.WizardClassicStyle.Automatic;
            this.wizardControl1.Location = new System.Drawing.Point(0, 0);
            this.wizardControl1.Name = "wizardControl1";
            this.wizardControl1.Pages.Add(this.wizardPage1);
            this.wizardControl1.Size = new System.Drawing.Size(524, 518);
            this.wizardControl1.TabIndex = 1;
            this.wizardControl1.Title = "文章作成ソフトウェアが使用できるか確認";
            // 
            // wizardPage1
            // 
            this.wizardPage1.AllowBack = false;
            this.wizardPage1.AllowCancel = false;
            this.wizardPage1.Controls.Add(this.kryptonPanel1);
            this.wizardPage1.Controls.Add(this.kryptonLabel1);
            this.wizardPage1.Name = "wizardPage1";
            this.wizardPage1.Size = new System.Drawing.Size(477, 364);
            this.wizardPage1.TabIndex = 0;
            this.wizardPage1.Text = "文章作成ソフトウェアを使用できることを確認できませんでした";
            // 
            // kryptonPanel1
            // 
            this.kryptonPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.kryptonPanel1.Location = new System.Drawing.Point(0, 174);
            this.kryptonPanel1.Name = "kryptonPanel1";
            this.kryptonPanel1.PanelBackStyle = ComponentFactory.Krypton.Toolkit.PaletteBackStyle.GridHeaderColumnCustom1;
            this.kryptonPanel1.Size = new System.Drawing.Size(477, 190);
            this.kryptonPanel1.StateCommon.Color1 = System.Drawing.Color.Transparent;
            this.kryptonPanel1.TabIndex = 3;
            // 
            // kryptonLabel1
            // 
            this.kryptonLabel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.kryptonLabel1.Location = new System.Drawing.Point(0, 0);
            this.kryptonLabel1.Name = "kryptonLabel1";
            this.kryptonLabel1.Size = new System.Drawing.Size(477, 84);
            this.kryptonLabel1.TabIndex = 2;
            this.kryptonLabel1.Values.Text = "このダイアログが表示された場合、文書作成ソフトウェアがインストールまたは使用できることを\r\n確認きなかったため文書作成ソフトウェアと連携ができない可能性があります" +
    "。ご注意ください\r\n。\r\n\r\n作業を続行するには「完了」をクリックしてください。";
            // 
            // WindowForWordIntegrationError
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 518);
            this.Controls.Add(this.wizardControl1);
            this.Controls.Add(this.wizardPageContainer1);
            this.Name = "WindowForWordIntegrationError";
            this.Text = "WindowForWordIntegrationError";
            ((System.ComponentModel.ISupportInitialize)(this.wizardPageContainer1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wizardControl1)).EndInit();
            this.wizardPage1.ResumeLayout(false);
            this.wizardPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPanel1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AeroWizard.WizardPageContainer wizardPageContainer1;
        private AeroWizard.WizardControl wizardControl1;
        private AeroWizard.WizardPage wizardPage1;
        private ComponentFactory.Krypton.Toolkit.KryptonLabel kryptonLabel1;
        private ComponentFactory.Krypton.Toolkit.KryptonPanel kryptonPanel1;
    }
}