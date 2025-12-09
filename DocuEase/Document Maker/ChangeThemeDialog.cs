using ComponentFactory.Krypton.Toolkit;
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
    public partial class ChangeThemeDialog :ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public ChangeThemeDialog()
        {
            InitializeComponent();
        }

        private void kryptonLabel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel1_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel2_Paint(object sender, PaintEventArgs e)
        {

        }

        //テーマ選択処理
        //Office2007
        private void kryptonRadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            //テーマカラーを確認してから変更
            //青
            if (kryptonRadioButton4.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //銀色
            else if (kryptonRadioButton3.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //黒
            else if (kryptonRadioButton5.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }

        }

        //Office2010
        private void kryptonRadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            //テーマカラーを確認してから変更
            //青
            if (kryptonRadioButton4.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //銀色
            else if (kryptonRadioButton3.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //黒
            else if (kryptonRadioButton5.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }


        //テーマカラー変更処理
        //青
        private void kryptonRadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            //テーマの状態を確認してから変更
            //Office2007の場合
            if (kryptonRadioButton1.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2010の場合
            else if (kryptonRadioButton2.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
        }

        //銀色
        private void kryptonRadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            //テーマの状態を確認してから変更
            //Office2007の場合
            if (kryptonRadioButton1.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2010の場合
            else if (kryptonRadioButton2.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
        }

        //黒
        private void kryptonRadioButton5_CheckedChanged(object sender, EventArgs e)
        {
            //テーマの状態を確認してから変更
            //Office2007の場合
            if (kryptonRadioButton1.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010の場合
            else if (kryptonRadioButton2.Checked == true)
            {
                SamplekryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }

        //現在のテーマ取得処理(リセットでも使うのでLoadイベントとは分割する)
        public void GetTheme()
        {
            //テーマ取得処理
            //Office2007青の場合
            if (Properties.Settings.Default.Theme == "Office2007Blue")
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;

                kryptonRadioButton4.Checked = true;
                kryptonRadioButton3.Checked = false;
                kryptonRadioButton5.Checked = false;
            }
            //Office2007銀色の場合
            else if (Properties.Settings.Default.Theme == "Office2007Silver")
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;

                kryptonRadioButton4.Checked = false;
                kryptonRadioButton3.Checked = true;
                kryptonRadioButton5.Checked = false;
            }
            //Office2007黒の場合
            else if (Properties.Settings.Default.Theme == "Office2007Black")
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;

                kryptonRadioButton4.Checked = false;
                kryptonRadioButton3.Checked = false;
                kryptonRadioButton5.Checked = true;
            }
            //Office2010青の場合
            else if (Properties.Settings.Default.Theme == "Office2010Blue")
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;

                kryptonRadioButton4.Checked = true;
                kryptonRadioButton3.Checked = false;
                kryptonRadioButton5.Checked = false;
            }
            //Office2010銀色の場合
            else if (Properties.Settings.Default.Theme == "Office2010Silver")
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;

                kryptonRadioButton4.Checked = false;
                kryptonRadioButton3.Checked = true;
                kryptonRadioButton5.Checked = false;
            }
            //Office2010黒の場合
            else if (Properties.Settings.Default.Theme == "Office2010Black")
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;

                kryptonRadioButton4.Checked = false;
                kryptonRadioButton3.Checked = false;
                kryptonRadioButton5.Checked = true;
            }


            //リボンシェイプ設定
            if (Properties.Settings.Default.UseOffice2007RibbonShape == true)
            {
                kryptonCheckBox2.Checked = true;

            }
            else if (Properties.Settings.Default.UseOffice2007RibbonShape == false)
            {
                kryptonCheckBox2.Checked = false;
            }

            //リボンかメニューバー設定
            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                kryptonRadioButton6.Checked = true;
                kryptonRadioButton8.Checked = false;
            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                kryptonRadioButton6.Checked = false;
                kryptonRadioButton8.Checked = true;
            }
        }

        private void ChangeThemeDialog_Load(object sender, EventArgs e)
        {

            //フォームロード時アプリの外観に合わせて変更する
            //Office2007青色
            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette2.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette2.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette2.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette2.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette2.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette2.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            //現在のテーマを適用
            GetTheme();

        }


        public string SelectedTheme { get; set; }
        public bool UseOffice2007RibbonMenuAndQAT { get; set; }

        public string RibbonOrMenuBar { get; set; }
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if(this.DialogResult == DialogResult.OK)
            {
                //選択されたテーマを保存
                SelectedTheme = SamplekryptonPalette1.BasePaletteMode.ToString();
                //リボンシェイプ設定を保存
                if(kryptonCheckBox2.Checked == true)
                {
                    UseOffice2007RibbonMenuAndQAT = true;
                }
                else if(kryptonCheckBox2.Checked == false)
                {
                    UseOffice2007RibbonMenuAndQAT = false;
                }

                //リボンかメニューバー設定を保存
                if(kryptonRadioButton6.Checked == true)
                {
                    RibbonOrMenuBar = "Ribbon";
                }
                else if(kryptonRadioButton8.Checked == true)
                {
                    RibbonOrMenuBar = "MenuBar";
                }

            }
        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            //現在のテーマを適用
            GetTheme();
        }

        private void kryptonPage1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonLabel9_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel10_Click(object sender, EventArgs e)
        {
            if(kryptonCheckBox2.Checked == false)
            {
                kryptonCheckBox2.Checked = true;
            }
            else if (kryptonCheckBox2.Checked == true)
            {
                kryptonCheckBox2.Checked = false;
            }
        }

        private void kryptonLabel8_Click(object sender, EventArgs e)
        {
            kryptonRadioButton6.Checked = true;
            kryptonRadioButton8.Checked = false;
        }

        private void kryptonLabel9_Click(object sender, EventArgs e)
        {
            kryptonRadioButton6.Checked = false;
            kryptonRadioButton8.Checked = true;
        }

        private void kryptonLabel10_MouseEnter(object sender, EventArgs e)
        {
        }
    }
}
