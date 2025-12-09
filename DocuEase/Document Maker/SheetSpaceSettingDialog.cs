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
    public partial class SheetSpaceSettingDialog : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public SheetSpaceSettingDialog()
        {
            InitializeComponent();
        }

        //上
        public int TopMargin {  get; set; }
        //下
        public int ButtomMargin {  get; set; }
        //左
        public int LeftMargin { get; set; }
        //右
        public int RightMargin { get; set; }

        private void SheetSpaceSettingDialog_Load(object sender, EventArgs e)
        {
            //余白表示を初期化
            panel15.Height = 0;
            panel16.Height = 0;
            panel13.Height = 0;
            panel14.Width = 0;

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

            kryptonNumericUpDown4.Value = TopMargin;
            kryptonNumericUpDown7.Value = ButtomMargin;
            kryptonNumericUpDown5.Value = RightMargin; 
            kryptonNumericUpDown6.Value = LeftMargin;

        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            TopMargin = (int)kryptonNumericUpDown4.Value;
            ButtomMargin = (int)kryptonNumericUpDown7.Value;
            LeftMargin = (int)kryptonNumericUpDown5.Value;
            RightMargin = (int)kryptonNumericUpDown6.Value;
        }

        private void kryptonNumericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            panel15.Height = (int)kryptonNumericUpDown4.Value / 10;
        }

        private void kryptonNumericUpDown7_ValueChanged(object sender, EventArgs e)
        {
            panel16.Height = (int)kryptonNumericUpDown7.Value / 10;
        }

        private void kryptonNumericUpDown5_ValueChanged(object sender, EventArgs e)
        {
            panel13.Width = (int)kryptonNumericUpDown5.Value / 10;
        }

        private void kryptonNumericUpDown6_ValueChanged(object sender, EventArgs e)
        {

            panel14.Width = (int)kryptonNumericUpDown6.Value / 10;
        }

        private void SheetSpaceSettingDialog_Shown(object sender, EventArgs e)
        {

        }


        //テンプレート
        //標準
        private void kryptonContextMenuItem31_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 138;
            kryptonNumericUpDown7.Value = 118;
            kryptonNumericUpDown5.Value = 118;
            kryptonNumericUpDown6.Value = 118;
        }

        //狭い
        private void kryptonContextMenuItem32_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 50;
            kryptonNumericUpDown7.Value = 50;
            kryptonNumericUpDown5.Value = 50;
            kryptonNumericUpDown6.Value = 50;
        }

        //やや狭い
        private void kryptonContextMenuItem33_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 100;
            kryptonNumericUpDown7.Value = 100;
            kryptonNumericUpDown5.Value = 75;
            kryptonNumericUpDown6.Value = 75;
        }

        //広い
        private void kryptonContextMenuItem34_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = 100;
            kryptonNumericUpDown7.Value = 100;
            kryptonNumericUpDown5.Value = 200;
            kryptonNumericUpDown6.Value = 200;
        }

        //リセット
        //上
        private void kryptonContextMenuItem22_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = TopMargin;
        }

        //下
        private void kryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown7.Value = ButtomMargin;
        }

        //右
        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown5.Value = LeftMargin;
        }

        //左
        private void kryptonContextMenuItem25_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown6.Value = RightMargin;
        }

        //すべて
        private void kryptonContextMenuItem26_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = TopMargin;
            kryptonNumericUpDown7.Value = ButtomMargin;
            kryptonNumericUpDown5.Value = LeftMargin;
            kryptonNumericUpDown6.Value = RightMargin;
        }

        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = Properties.Settings.Default.Space_Top;
            kryptonNumericUpDown7.Value = Properties.Settings.Default.Space_Buttom;
            kryptonNumericUpDown5.Value = Properties.Settings.Default.Space_Left;
            kryptonNumericUpDown6.Value = Properties.Settings.Default.Space_Right;
        }
    }
}
