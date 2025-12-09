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
    public partial class AddWebLinkDialog : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public AddWebLinkDialog()
        {
            InitializeComponent();
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            try
            {
                //文頭にhttpsもしくはhttpがついてある場合そのまま通し、そうではない場合、httpsを付けてから通す
                if (kryptonTextBox1.Text.Contains("https://") == true | kryptonTextBox1.Text.Contains("http://") == true)
                {
                    webView21.Source = new Uri(kryptonTextBox1.Text);
                    kryptonTextBox2.Text = webView21.Source.ToString();
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }
                else
                {
                    webView21.Source = new Uri("https://" + kryptonTextBox1.Text);
                    kryptonTextBox2.Text = webView21.Source.ToString();
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }
            }
            catch
            {
                kryptonLabel5.Visible = true;
            }

        }

        private void webView21_ContentLoading(object sender, Microsoft.Web.WebView2.Core.CoreWebView2ContentLoadingEventArgs e)
        {
            if (webView21.Source != new Uri("about:blank"))
            {
                if (kryptonTextBox1.Text.Contains("https://") == true | kryptonTextBox1.Text.Contains("http://") == true)
                {
                    webView21.Source = new Uri(kryptonTextBox1.Text);
                    kryptonTextBox2.Text = webView21.Source.ToString();
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }
                else
                {
                    webView21.Source = new Uri("https://" + kryptonTextBox1.Text);
                    kryptonTextBox2.Text = webView21.Source.ToString();
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }

            }
            else
            {
                buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
            }

        }

        private void buttonSpecAny1_Click(object sender, EventArgs e)
        {
            kryptonTextBox1.Text = string.Empty;
            //空白ページに切り替える
            webView21.Source = new Uri("about:blank");
            kryptonTextBox2.Text = webView21.Source.ToString();
            buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
        }

        private void webView21_ContentLoading(object sender, EventArgs e)
        {
            if (webView21.Source != new Uri("about:blank"))
            {
                if (kryptonTextBox1.Text.Contains("https://") == true | kryptonTextBox1.Text.Contains("http://") == true)
                {
                    webView21.Source = new Uri(kryptonTextBox1.Text);
                    kryptonTextBox2.Text = webView21.Source.ToString();
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }
                else
                {
                    webView21.Source = new Uri("https://" + kryptonTextBox1.Text);
                    kryptonTextBox2.Text = webView21.Source.ToString();
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }

            }
            else
            {
                buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
            }
        }

        private void AddWebLinkDialog_Load(object sender, EventArgs e)
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

            kryptonTextBox2.Text = webView21.Source.ToString();
        }

        private void AddWebLinkDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            webView21.Dispose();
        }

        private void buttonSpecAny4_Click(object sender, EventArgs e)
        {
            if(webView21.Source != new Uri("about:blank"))
            {
                System.Diagnostics.Process.Start(webView21.Source.ToString());
            }
            
        }

        private void kryptonTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                if (webView21.Source != new Uri("about:blank"))
                {
                    if(kryptonTextBox1.Text.Contains("https://") == true| kryptonTextBox1.Text.Contains("http://") == true)
                    {
                        webView21.Source = new Uri(kryptonTextBox1.Text);
                        kryptonTextBox2.Text = webView21.Source.ToString();
                        buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                    }
                    else
                    {
                        webView21.Source = new Uri("https://"+kryptonTextBox1.Text);
                        kryptonTextBox2.Text = webView21.Source.ToString();
                        buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                    }

                }
                else
                {
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
                }
            }
        }

        public string WebLink { get; set; }
        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox1.Text.Contains("https://") == true | kryptonTextBox1.Text.Contains("http://") == true)
            {
                try
                {
                    Uri uri = new Uri(kryptonTextBox1.Text);
                    WebLink = kryptonTextBox1.Text;
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
                }
                catch
                {
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }

            }
            else
            {
                try
                {
                    Uri uri = new Uri("https://" + kryptonTextBox1.Text);
                    WebLink = "https://" + kryptonTextBox1.Text;
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
                }
                catch
                {
                    buttonSpecAny4.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                }
            }
            
        }
    }
}
