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
    public partial class ShareDialog :　ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public string ShareContent { get; set; }
        public string ShareTitle { get; set; }
        public ShareDialog()
        {
            InitializeComponent();
            
        }

        private void kryptonCommandLinkButton1_Click(object sender, EventArgs e)
        {
            ShareContent = "MicrosoftOutlook";
            this.Close();
        }

        private void kryptonCommandLinkButton2_Click(object sender, EventArgs e)
        {
            ShareContent = "MicrosoftTeams";
            this.Close();
        }

        private void kryptonCommandLinkButton3_Click(object sender, EventArgs e)
        {
            ShareContent = "Slack";
            this.Close();
        }

        private void ShareDialog_Load(object sender, EventArgs e)
        {
            kryptonLabel2.Text = ShareTitle;
            //Office2007青色
            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonCommandLinkButton3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }
    }
}
