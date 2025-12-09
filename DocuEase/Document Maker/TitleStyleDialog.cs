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
    public partial class TitleStyleDialog : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public string LoadTitleStyle {  get; set; }

        public TitleStyleDialog()
        {
            InitializeComponent();
        }


        private void TitleStyleDialog_Load(object sender, EventArgs e)
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

        }

        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = true;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;
        }

        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = true;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;
        }

        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = true;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;
        }

        private void kryptonCheckButton4_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = true;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;
        }

        private void kryptonCheckButton5_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = true;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;
        }

        private void kryptonCheckButton6_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = true;
            kryptonCheckButton7.Checked = false;
        }

        private void kryptonCheckButton7_Click(object sender, EventArgs e)
        {
            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = false;
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = true;
        }

        public string SetTileStyle {  get; set; }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if(kryptonCheckButton1.Checked == true)
            {
                SetTileStyle = "Defalt";
            }
            else if(kryptonCheckButton2.Checked == true)
            {
                SetTileStyle = "Headline";
            }
            else if (kryptonCheckButton3.Checked == true)
            {
                SetTileStyle = "Modern";
            }
            else if (kryptonCheckButton4.Checked == true)
            {
                SetTileStyle = "ModernBold";
            }
            else if (kryptonCheckButton5.Checked == true)
            {
                SetTileStyle = "Note";
            }
            else if (kryptonCheckButton6.Checked == true)
            {
                SetTileStyle = "Emphasis";
            }
            else if (kryptonCheckButton7.Checked == true)
            {
                SetTileStyle = "Cancel";
            }
        }
    }
}
