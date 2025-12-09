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
    public partial class AddressDeleteDone : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public AddressDeleteDone()
        {
            InitializeComponent();
        }

        private void AddressDeleteDone_Load(object sender, EventArgs e)
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
            //Office2007ブラック
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

        private void AddressDeleteDone_Shown(object sender, EventArgs e)
        {
            System.Media.SystemSounds.Asterisk.Play();
        }
    }
}
