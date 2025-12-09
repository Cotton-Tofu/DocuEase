using ComponentFactory.Krypton.Ribbon;
using Krypton.Toolkit;
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
    public partial class SplashWindow : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public SplashWindow()
        {
            InitializeComponent();


        }

        private void SplashWindow_Load(object sender, EventArgs e)
        {
            kryptonLabel4.Text = System.Windows.Forms.Application.ProductVersion;

            //2007
            if (Properties.Settings.Default.Theme == "Office2007Blue")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            else if (Properties.Settings.Default.Theme == "Office2007Silver")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (Properties.Settings.Default.Theme == "Office2007Black")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLabel1.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.TitleControl;
                kryptonLabel2.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl;
                kryptonLabel3.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl;
                kryptonLabel4.LabelStyle = ComponentFactory.Krypton.Toolkit.LabelStyle.NormalControl;
            }
            //2010
            else if (Properties.Settings.Default.Theme == "Office2010Blue")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            else if (Properties.Settings.Default.Theme == "Office2010Silver")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            else if (Properties.Settings.Default.Theme == "Office2010Black")
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }
    }
}
