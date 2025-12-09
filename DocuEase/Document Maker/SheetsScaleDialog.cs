using ComponentFactory.Krypton.Ribbon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Document_Maker
{
    public partial class SheetsScaleDialog :ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public SheetsScaleDialog()
        {
            InitializeComponent();
        }

        

        private void kryptonTrackBar1_ValueChanged(object sender, EventArgs e)
        {
            //10の目盛りに合わせてサイズを+50上げる
            if (kryptonTrackBar1.Value == 0)
            {

                Form1.Form1Instance.Sheets_Sheet.Size = new Size(842, 999);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                //縮小ボタンを無効化
                Form1.Form1Instance.kryptonButton15.Enabled = false;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "10%";
                kryptonLabel42.Text = "10%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "10";
            }
            else if (kryptonTrackBar1.Value == 1)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(892, 1049);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;

                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "20%";
                kryptonLabel42.Text = "20%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "20";
            }
            else if (kryptonTrackBar1.Value == 2)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(942, 1099);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;
                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "30%";
                kryptonLabel42.Text = "30%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "30";
            }
            else if (kryptonTrackBar1.Value == 3)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(992, 1149);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;
                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "40%";
                kryptonLabel42.Text = "40%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "40";
            }
            else if (kryptonTrackBar1.Value == 4)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1042, 1199);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;
                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "50%";
                kryptonLabel42.Text = "50%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "50";
            }
            else if (kryptonTrackBar1.Value == 5)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1092, 1249);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;
                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "60%";
                kryptonLabel42.Text = "60%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "60";
            }
            else if (kryptonTrackBar1.Value == 6)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1142, 1299);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;
                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "70%";
                kryptonLabel42.Text = "70%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "70";
            }
            else if (kryptonTrackBar1.Value == 7)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1192, 1349);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;

                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;
                
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "80%";
                kryptonLabel42.Text = "80%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "80";
            }
            else if (kryptonTrackBar1.Value == 8)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1242, 1399);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;

                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "90%";
                kryptonLabel42.Text = "90%";

                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "90";
            }
            else if (kryptonTrackBar1.Value == 9)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1292, 1449);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;

                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = true;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                Form1.Form1Instance.kryptonLabel42.Text = "100%";
                kryptonLabel42.Text = "100%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "100";
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                Form1.Form1Instance.Sheets_Sheet.Size = new Size(1342, 1499);

                Form1.Form1Instance.kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Form1.Form1Instance.Sheets_Sheet.Anchor = AnchorStyles.Top;

                Form1.Form1Instance.kryptonTrackBar1.Value = kryptonTrackBar1.Value;

                // 親コントロールのサイズを取得
                int parentWidth = Form1.Form1Instance.ClientSize.Width;
                int parentHeight = Form1.Form1Instance.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Form1.Form1Instance.Sheets_Sheet.Width;
                int panelHeight = Form1.Form1Instance.Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Form1.Form1Instance.Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Form1.Form1Instance.Sheets_Sheet.Top = 59;

                Form1.Form1Instance.kryptonButton15.Enabled = true;
                Form1.Form1Instance.kryptonButton14.Enabled = false;

                kryptonButton15.Enabled = true;
                //拡大ボタンを無効化
                kryptonButton14.Enabled = false;

                Form1.Form1Instance.kryptonLabel42.Text = "110%";
                kryptonLabel42.Text = "110%";
                Form1.Form1Instance.kryptonRibbonGroupComboBox1.Text = "110";
            }
        }

        private void SheetsScaleDialog_Load(object sender, EventArgs e)
        {
            
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(kryptonCheckBox1.Checked == true)
            {
                this.TopMost = true;
            }
            else if (kryptonCheckBox1.Checked == false)
            { 
                this.TopMost = false; 
            }
        }

        private void SheetsScaleDialog_Paint(object sender, PaintEventArgs e)
        {
            //Office2007青色
            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            //初期化
            kryptonTrackBar1.Value = Form1.Form1Instance.kryptonTrackBar1.Value;
        }

        private void SheetsScaleDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();

        }

        private void kryptonButton14_Click(object sender, EventArgs e)
        {
            if (kryptonTrackBar1.Value == kryptonTrackBar1.Value)
            {
                kryptonTrackBar1.Value = kryptonTrackBar1.Value + 1;
            }
        }

        private void kryptonButton15_Click(object sender, EventArgs e)
        {
            if (kryptonTrackBar1.Value == kryptonTrackBar1.Value)
            {
                kryptonTrackBar1.Value = kryptonTrackBar1.Value - 1;
            }
        }
    }
}
