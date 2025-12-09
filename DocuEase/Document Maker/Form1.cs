using ComponentFactory.Krypton.Navigator;
using ComponentFactory.Krypton.Ribbon;
using ComponentFactory.Krypton.Toolkit;
using FluentTransitions;
using Krypton.Toolkit;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using Microsoft.Toolkit.Uwp.Notifications;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Foundation.Metadata;
using Windows.Security.Credentials.UI;
using Windows.UI.Xaml.Documents;
using Windows.UI.Xaml.Shapes;

namespace Document_Maker
{

    public partial class Form1 : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        SplashWindow splashWindow = new SplashWindow();

        public Form1()
        {
            InitializeComponent();
            TaskAsync();



        }

        async System.Threading.Tasks.Task TaskAsync()
        {
            splashWindow.Show();
        }


        TreeNode treeNode1 = new TreeNode();


        private void kryptonRibbonGroupCheckBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        //下準備
        public void SetTheme()
        {

            //AppButtonAndQAT
            if (Properties.Settings.Default.UseOffice2007RibbonShape == true)
            {
                kryptonRibbon.StateCommon.RibbonGeneral.RibbonShape = ComponentFactory.Krypton.Toolkit.PaletteRibbonShape.Office2007;
                kryptonContextMenuItem93.Checked = true;
            }
            else if (Properties.Settings.Default.UseOffice2007RibbonShape == false)
            {
                kryptonRibbon.StateCommon.RibbonGeneral.RibbonShape = ComponentFactory.Krypton.Toolkit.PaletteRibbonShape.Inherit;
                kryptonContextMenuItem93.Checked = false;
            }

            //2007
            if (Properties.Settings.Default.Theme == "Office2007Blue")
            {
                //テーマの設定
                of2007.Checked = true;
                of2010.Checked = false;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = true;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(191, 219, 255);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(191, 219, 255);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007青色
                this.BackColor = Color.FromArgb(191, 219, 255);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(191, 219, 255);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            else if (Properties.Settings.Default.Theme == "Office2007Silver")
            {
                //テーマの設定
                of2007.Checked = true;
                of2010.Checked = false;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = true;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(208, 212, 221);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(208, 212, 221);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007銀色
                this.BackColor = Color.FromArgb(208, 212, 221);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(208, 212, 221);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (Properties.Settings.Default.Theme == "Office2007Black")
            {
                //テーマの設定
                of2007.Checked = true;
                of2010.Checked = false;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = true;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(83, 83, 83);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(83, 83, 83);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007黒色
                this.BackColor = Color.FromArgb(83, 83, 83);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(83, 83, 83);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //2010
            else if (Properties.Settings.Default.Theme == "Office2010Blue")
            {
                //テーマの設定
                of2007.Checked = false;
                of2010.Checked = true;

                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = true;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(187, 206, 230);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(187, 206, 230);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010青色
                this.BackColor = Color.FromArgb(187, 206, 230);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(187, 206, 230);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            else if (Properties.Settings.Default.Theme == "Office2010Silver")
            {
                //テーマの設定
                of2007.Checked = false;
                of2010.Checked = true;

                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = true;
                kryptonContextMenuRadioButton3.Checked = false;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(227, 230, 232);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(227, 230, 232);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010銀色
                this.BackColor = Color.FromArgb(227, 230, 232);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(227, 230, 232);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            else if (Properties.Settings.Default.Theme == "Office2010Black")
            {
                //テーマの設定
                of2007.Checked = false;
                of2010.Checked = true;
                //ラジオボタンの設定
                kryptonContextMenuRadioButton1.Checked = false;
                kryptonContextMenuRadioButton2.Checked = false;
                kryptonContextMenuRadioButton3.Checked = true;

                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(113, 113, 113);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(113, 113, 113);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010黒色
                this.BackColor = Color.FromArgb(113, 113, 113);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(113, 113, 113);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }

        public void SetQAT()
        {
            if (Properties.Settings.Default.ShowQATLocation == 0)
            {
                kryptonRibbon.QATLocation = QATLocation.Above;
            }
            else if (Properties.Settings.Default.ShowQATLocation == 1)
            {
                kryptonRibbon.QATLocation = QATLocation.Below;
            }

            if (Properties.Settings.Default.QAT1_Visible == true)
            {
                kryptonRibbonQATButton1.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton1.Visible = false;
            }

            if (Properties.Settings.Default.QAT2_Visible == true)
            {
                kryptonRibbonQATButton2.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton2.Visible = false;
            }

            if (Properties.Settings.Default.QAT3_Visible == true)
            {
                kryptonRibbonQATButton3.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton3.Visible = false;
            }

            if (Properties.Settings.Default.QAT4_Visible == true)
            {
                kryptonRibbonQATButton4.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton4.Visible = false;
            }

            if (Properties.Settings.Default.QAT5_Visible == true)
            {
                kryptonRibbonQATButton5.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton5.Visible = false;
            }

            if (Properties.Settings.Default.QAT6_Visible == true)
            {
                kryptonRibbonQATButton6.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton6.Visible = false;
            }

            if (Properties.Settings.Default.QAT7_Visible == true)
            {
                kryptonRibbonQATButton7.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton7.Visible = false;
            }

            if (Properties.Settings.Default.QAT8_Visible == true)
            {
                kryptonRibbonQATButton8.Visible = true;
            }
            else
            {
                kryptonRibbonQATButton8.Visible = false;
            }
        }

        public void SetSheetSpace()
        {
            Sheets_TopPanel.Height = Properties.Settings.Default.Space_Top;
            Sheets_ButtomPanel.Height = Properties.Settings.Default.Space_Buttom;
            Sheets_LeftPanel.Width = Properties.Settings.Default.Space_Left;
            Sheets_RightPanel.Width = Properties.Settings.Default.Space_Right;

            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = Properties.Settings.Default.Space_Top;
            kryptonRibbonGroupNumericUpDown1.Value = Properties.Settings.Default.Space_Buttom;
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = Properties.Settings.Default.Space_Left;
            kryptonRibbonGroupNumericUpDown2.Value = Properties.Settings.Default.Space_Right;
        }

        public void SetSheetText()
        {
            //発行元部署
            kryptonTextBox11.Text = Properties.Settings.Default.SendingDepartment;
            //宛先会社
            kryptonTextBox1.Text = Properties.Settings.Default.To_CompanyOrOrganizationName;
            //宛先肩書
            kryptonComboBox10.Text = Properties.Settings.Default.To_Title;
            //宛先氏名
            kryptonTextBox2.Text = Properties.Settings.Default.To_Name;

            //発信者会社
            kryptonTextBox3.Text = Properties.Settings.Default.Caller_CompanyOrOrganizationName;
            //発信者所在地
            kryptonTextBox4.Text = Properties.Settings.Default.Caller_Location;
            //発信者建物名
            kryptonTextBox5.Text = Properties.Settings.Default.Caller_BuildingName;
            //発信者階数
            kryptonNumericUpDown2.Value = Properties.Settings.Default.Caller_FloorNumber;
            //発信者肩書
            kryptonComboBox9.Text = Properties.Settings.Default.Caller_Title;
            //発信者氏名
            kryptonTextBox6.Text = Properties.Settings.Default.Caller_Name;
            //メールアドレス
            kryptonTextBox7.Text = Properties.Settings.Default.Caller_MailAddress_User;
            kryptonComboBox8.Text = Properties.Settings.Default.Caller_MailAddress_Domain;
            //電話番号1
            kryptonComboBox6.Text = Properties.Settings.Default.Caller_PhoneNumber1;
            kryptonTextBox14.Text = Properties.Settings.Default.Caller_PhoneNumber2;
            kryptonTextBox8.Text = Properties.Settings.Default.Caller_PhoneNumber3;
            //Fax番号
            kryptonComboBox7.Text = Properties.Settings.Default.Caller_FaxNumber1;
            kryptonTextBox9.Text = Properties.Settings.Default.Caller_FaxNumber2;
            kryptonTextBox15.Text = Properties.Settings.Default.Caller_FaxNumber3;


        }

        public void RunAppTask()
        {

            if (Properties.Settings.Default.ShowApplicationTask == 0)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = true;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 1)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;
                kryptonRadioButton3.Checked = false;
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonTrackBar1.Enabled = false;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = false;

                kryptonRibbon.Enabled = false;
                kryptonRibbon.MinimizedMode = true;
                kryptonPage9.Visible = true;
                kryptonNavigator_Workbench.SelectedPage = kryptonPage9;
                kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.Panel;
                this.Text = "テンプレート - DoQuick Designer";

                kryptonLabel7.Enabled = false;
                kryptonCheckButton1.Enabled = false;
                kryptonCheckButton2.Enabled = false;
                kryptonLabel1.Enabled = false;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 2)
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = false;
                ShowDcw();

            }
        }

        public void ShowDcw()
        {
            DCW dCW = new DCW();

            Properties.Settings.Default.dCW_TopSpace = Sheets_TopPanel.Height;
            Properties.Settings.Default.dCW_ButtomSpace = Sheets_ButtomPanel.Height;
            Properties.Settings.Default.dCW_LeftSpace = Sheets_LeftPanel.Width;
            Properties.Settings.Default.dCW_RightSpace = Sheets_RightPanel.Width;
            Properties.Settings.Default.Save();

            // Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                dCW.BackColor = Color.FromArgb(191, 219, 255);
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                dCW.BackColor = Color.FromArgb(208, 212, 221);
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                dCW.BackColor = Color.FromArgb(83, 83, 83);
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                dCW.BackColor = Color.FromArgb(187, 206, 230);
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                dCW.BackColor = Color.FromArgb(227, 230, 232);
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                dCW.BackColor = Color.FromArgb(113, 113, 113);
            }

            dCW.StartPosition = FormStartPosition.CenterScreen;
            dCW.ShowDialog();

            //trueの場合のみ実行
            //ウィザードの「完了」ボタンをクリックしたときに実行
            if (dCW.IsWizardFinished == true)
            {
                FontReset();
                //発行番号
                if (dCW.NoIssueNumber == false)
                {
                    kryptonCheckBox3.Checked = false;
                    kryptonTextBox11.Text = dCW.IssueNumber_Publisher;
                    kryptonNumericUpDown1.Value = dCW.IssueNumber;
                }
                else
                {
                    kryptonCheckBox3.Checked = true;
                }

                //日付
                if (dCW.NoDate == false)
                {
                    kryptonCheckBox2.Checked = false;
                    kryptonDateTimePicker1.Value = dCW.Date;
                    if (dCW.UseEraName == true)
                    {
                        kryptonCheckBox1.Checked = true;
                    }
                    else
                    {
                        kryptonCheckBox1.Checked = false;
                    }
                }
                else
                {
                    kryptonCheckBox2.Checked = true;
                }

                //発信者
                kryptonTextBox1.Text = dCW.AdCompany;
                kryptonComboBox10.Text = dCW.AdTitle;
                kryptonTextBox2.Text = dCW.AdName;

                kryptonTextBox3.Text = dCW.CaCampany;
                kryptonTextBox4.Text = dCW.CaLocation;
                kryptonTextBox5.Text = dCW.CaBuildingName;
                kryptonNumericUpDown2.Value = dCW.CaFloorNumber;
                kryptonComboBox9.Text = dCW.CaTitle;
                kryptonTextBox6.Text = dCW.CaName;
                kryptonTextBox7.Text = dCW.CaMailAddress;
                kryptonComboBox8.Text = dCW.CaMailAddress_Domain;
                //電話番号
                kryptonComboBox6.Text = dCW.CaPhoneNumber1;
                kryptonTextBox14.Text = dCW.CaPhoneNumber2;
                kryptonTextBox8.Text = dCW.CaPhoneNumber3;
                kryptonComboBox7.Text = dCW.CaFaxNumber1;
                kryptonTextBox9.Text = dCW.CaFaxNumber2;
                kryptonTextBox15.Text = dCW.CaFaxNumber3;

                //表題
                kryptonTextBox10.Text = dCW.title;
                kryptonTextBox10.StateCommon.Content.Color1 = dCW.titleColor;
                Sheets_TitleButton.ForeColor = dCW.titleColor;

                //表題のフォント
                kryptonRibbonGroupComboBox_Font.Text = dCW.ftName;
                kryptonRibbonGroupComboBox_FontSize.Text = dCW.ftSize.ToString();
                kryptonRibbonColorButton_TextColor.SelectedColor = dCW.titleColor;


                if (dCW.titleBold == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Bold);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;

                    kryptonRibbonButton_Bold.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Bold.Checked = false;
                }

                if (dCW.titleItalic == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Italic);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Italic.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Italic.Checked = false;
                }

                if (dCW.titleUnderline == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Underline);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem15.Checked = true;
                }
                if (dCW.titleUnderline == false)
                {
                    kryptonContextMenuItem15.Checked = false;
                }

                if (dCW.titleStrikeout == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Strikeout);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem16.Checked = true;
                }
                if (dCW.titleStrikeout == false)
                {
                    kryptonContextMenuItem16.Checked = false;
                }

                //あいさつ文
                //月
                kryptonComboBox1.Text = dCW.UseSourouBunDate;
                //頭語
                kryptonComboBox2.Text = dCW.acronym;
                //候文
                kryptonComboBox11.Text = dCW.souroubun;
                //前文
                kryptonComboBox3.Text = dCW.PreviousText;
                //感謝のあいさつ
                kryptonComboBox4.Text = dCW.ThankYouGreeting;
                //結語
                kryptonComboBox5.Text = dCW.Conclusion;

                //内容
                kryptonTextBox12.Text = dCW.Content;
                kryptonTextBox13.Text = dCW.Notetaking;
            }
            // falseの場合は何もしない
        }

        public void RunWordInstalled()
        {
            if (Properties.Settings.Default.IsAvailableDocumentCreationSoftware == true)
            {
                kryptonCheckBox4.Checked = true;
                if (IsWordInstalled() != true)
                {
                    WindowForWordIntegrationError windowForWordIntegrationError = new WindowForWordIntegrationError();
                    windowForWordIntegrationError.ShowDialog();
                }
            }
            else
            {
                kryptonCheckBox4.Checked = false;
            }
        }

        private bool IsWordInstalled()
        {
            const string wordRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE";

            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(wordRegistryKey))
            {
                return key != null;
            }
        }

        public void SetSettings()
        {

            if (Properties.Settings.Default.ShowApplicationTask == 0)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = true;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 1)
            {
                kryptonRadioButton1.Checked = false;
                kryptonRadioButton2.Checked = true;
                kryptonRadioButton3.Checked = false;
            }
            else if (Properties.Settings.Default.ShowApplicationTask == 2)
            {
                kryptonRadioButton1.Checked = true;
                kryptonRadioButton2.Checked = false;
                kryptonRadioButton3.Checked = false;

            }

            if (Properties.Settings.Default.IsAvailableDocumentCreationSoftware == true)
            {
                kryptonCheckBox4.Checked = true;
            }
            else
            {
                kryptonCheckBox4.Checked = false;
            }

            if (Properties.Settings.Default.IsWindowsStartUpRunForDCMK == true)
            {
                kryptonCheckBox5.Checked = true;
            }
            else
            {
                kryptonCheckBox5.Checked = false;
            }

            if (Properties.Settings.Default.IsUseEraName == true)
            {
                kryptonCheckBox1.Checked = true;
                kryptonCheckBox7.Checked = true;
            }
            else
            {
                kryptonCheckBox1.Checked = false;
                kryptonCheckBox7.Checked = false;
            }
            kryptonNumericUpDown4.Value = Properties.Settings.Default.Space_Top;
            kryptonNumericUpDown7.Value = Properties.Settings.Default.Space_Buttom;
            kryptonNumericUpDown5.Value = Properties.Settings.Default.Space_Left;
            kryptonNumericUpDown6.Value = Properties.Settings.Default.Space_Right;

            kryptonTextBox16.Text = Properties.Settings.Default.SendingDepartment;
            kryptonTextBox17.Text = Properties.Settings.Default.To_CompanyOrOrganizationName;
            kryptonComboBox12.Text = Properties.Settings.Default.To_Title;
            kryptonTextBox18.Text = Properties.Settings.Default.To_Name;
            kryptonTextBox19.Text = Properties.Settings.Default.Caller_CompanyOrOrganizationName;
            kryptonTextBox32.Text = Properties.Settings.Default.Caller_Location;
            kryptonTextBox20.Text = Properties.Settings.Default.Caller_BuildingName;
            kryptonNumericUpDown3.Value = Properties.Settings.Default.Caller_FloorNumber;
            kryptonComboBox13.Text = Properties.Settings.Default.Caller_Title;
            kryptonTextBox21.Text = Properties.Settings.Default.Caller_Name;
            kryptonTextBox22.Text = Properties.Settings.Default.Caller_MailAddress_User;
            kryptonComboBox14.Text = Properties.Settings.Default.Caller_MailAddress_Domain;
            kryptonComboBox15.Text = Properties.Settings.Default.Caller_PhoneNumber1;
            kryptonTextBox23.Text = Properties.Settings.Default.Caller_PhoneNumber2;
            kryptonTextBox24.Text = Properties.Settings.Default.Caller_PhoneNumber3;
            kryptonComboBox16.Text = Properties.Settings.Default.Caller_FaxNumber1;
            kryptonTextBox26.Text = Properties.Settings.Default.Caller_FaxNumber2;
            kryptonTextBox25.Text = Properties.Settings.Default.Caller_FaxNumber3;

        }
        #region TreeNodeの追加;

        TreeNode miniTreeNode1 = new TreeNode();
        TreeNode ultraMiniNode1 = new TreeNode();
        TreeNode ultraMiniNode2 = new TreeNode();
        TreeNode ultraMiniNode3 = new TreeNode();
        TreeNode ultraMiniNode4 = new TreeNode();
        TreeNode ultraMiniNode5 = new TreeNode();
        TreeNode miniTreeNode2 = new TreeNode();
        TreeNode ultraMiniNode6 = new TreeNode();
        TreeNode ultraMiniNode7 = new TreeNode();
        TreeNode ultraMiniNode8 = new TreeNode();
        TreeNode ultraMiniNode9 = new TreeNode();
        TreeNode ultraMiniNode10 = new TreeNode();
        TreeNode ultraMiniNode11 = new TreeNode();
        TreeNode treeNode2 = new TreeNode();
        TreeNode miniTreeNode3 = new TreeNode();
        TreeNode ultraMiniNode12 = new TreeNode();
        TreeNode ultraMiniNode13 = new TreeNode();
        TreeNode ultraMiniNode14 = new TreeNode();
        TreeNode ultraMiniNode15 = new TreeNode();
        TreeNode ultraMiniNode16 = new TreeNode();
        TreeNode ultraMiniNode17 = new TreeNode();
        TreeNode miniTreeNode4 = new TreeNode();
        TreeNode ultraMiniNode18 = new TreeNode();
        TreeNode ultraMiniNode19 = new TreeNode();
        TreeNode ultraMiniNode20 = new TreeNode();
        TreeNode hyperTreeNode1 = new TreeNode();
        TreeNode ultraMiniNode21 = new TreeNode();
        TreeNode treeNode3 = new TreeNode();
        TreeNode miniTreeNode22 = new TreeNode();
        TreeNode miniTreeNode23 = new TreeNode();
        TreeNode miniTreeNode24 = new TreeNode();
        TreeNode miniTreeNode25 = new TreeNode();
        TreeNode miniTreeNode26 = new TreeNode();
        TreeNode miniTreeNode27 = new TreeNode();
        TreeNode miniTreeNode28 = new TreeNode();
        TreeNode treeNode4 = new TreeNode();
        TreeNode miniTreeNode29 = new TreeNode();
        TreeNode miniTreeNode30 = new TreeNode();
        TreeNode miniTreeNode31 = new TreeNode();
        TreeNode miniTreeNode32 = new TreeNode();
        TreeNode miniTreeNode33 = new TreeNode();
        TreeNode miniTreeNode34 = new TreeNode();

        public void AddTreeNodes()
        {
            //TreeViewに各種ノードを追加する
            //ノード1

            treeNode1.Text = "取引文書";
            treeView1.Nodes.Add(treeNode1);
            //子ノード1

            miniTreeNode1.Text = "通常取引";
            treeNode1.Nodes.Add(miniTreeNode1);
            //孫ノード1

            ultraMiniNode1.Text = "注文書";
            miniTreeNode1.Nodes.Add(ultraMiniNode1);
            //孫ノード2

            ultraMiniNode2.Text = "承諾書";
            miniTreeNode1.Nodes.Add(ultraMiniNode2);
            //孫ノード3

            ultraMiniNode3.Text = "依頼文";
            miniTreeNode1.Nodes.Add(ultraMiniNode3);
            //孫ノード4

            ultraMiniNode4.Text = "照会文";
            miniTreeNode1.Nodes.Add(ultraMiniNode4);
            //孫ノード5

            ultraMiniNode5.Text = "回答文";
            miniTreeNode1.Nodes.Add(ultraMiniNode5);
            //子ノード2

            miniTreeNode2.Text = "例外的取引";
            treeNode1.Nodes.Add(miniTreeNode2);
            //孫ノード6

            ultraMiniNode6.Text = "催促文";
            miniTreeNode2.Nodes.Add(ultraMiniNode6);
            //孫ノード7

            ultraMiniNode7.Text = "断り文";
            miniTreeNode2.Nodes.Add(ultraMiniNode7);
            //孫ノード8

            ultraMiniNode8.Text = "交渉文";
            miniTreeNode2.Nodes.Add(ultraMiniNode8);
            //孫ノード9

            ultraMiniNode9.Text = "抗議文";
            miniTreeNode2.Nodes.Add(ultraMiniNode9);
            //孫ノード10

            ultraMiniNode10.Text = "お詫び文";
            miniTreeNode2.Nodes.Add(ultraMiniNode10);
            //孫ノード11

            ultraMiniNode11.Text = "取り消し文";
            miniTreeNode2.Nodes.Add(ultraMiniNode11);
            //ノード2

            treeNode2.Text = "社公文書";
            treeView1.Nodes.Add(treeNode2);
            //子ノード3

            miniTreeNode3.Text = "公的";
            treeNode2.Nodes.Add(miniTreeNode3);
            //孫ノード12

            ultraMiniNode12.Text = "あいさつ文";
            miniTreeNode3.Nodes.Add(ultraMiniNode12);
            //孫ノード13

            ultraMiniNode13.Text = "お祝い文";
            miniTreeNode3.Nodes.Add(ultraMiniNode13);
            //孫ノード14

            ultraMiniNode14.Text = "招待文";
            miniTreeNode3.Nodes.Add(ultraMiniNode14);
            //孫ノード15

            ultraMiniNode15.Text = "お礼文";
            miniTreeNode3.Nodes.Add(ultraMiniNode15);
            //孫ノード16

            ultraMiniNode16.Text = "案内文";
            miniTreeNode3.Nodes.Add(ultraMiniNode16);
            //孫ノード17

            ultraMiniNode17.Text = "通知文";
            miniTreeNode3.Nodes.Add(ultraMiniNode17);
            //子ノード4

            miniTreeNode4.Text = "私的";
            treeNode2.Nodes.Add(miniTreeNode4);
            //孫ノード18

            ultraMiniNode18.Text = "年賀文";
            miniTreeNode4.Nodes.Add(ultraMiniNode18);
            //孫ノード19

            ultraMiniNode19.Text = "季節のあいさつ文";
            miniTreeNode4.Nodes.Add(ultraMiniNode19);
            //孫ノード20

            ultraMiniNode20.Text = "見舞い文";
            miniTreeNode4.Nodes.Add(ultraMiniNode20);
            //赤子ノード1

            hyperTreeNode1.Text = "個人宛見舞い文";
            ultraMiniNode20.Nodes.Add(hyperTreeNode1);
            //孫ノード21

            ultraMiniNode21.Text = "お悔やみ文";
            miniTreeNode4.Nodes.Add(ultraMiniNode21);

            //ノード3

            treeNode3.Text = "連絡文書";
            treeView2.Nodes.Add(treeNode3);
            //孫ノード22

            miniTreeNode22.Text = "通達文";
            treeNode3.Nodes.Add(miniTreeNode22);
            //孫ノード23

            miniTreeNode23.Text = "指示文";
            treeNode3.Nodes.Add(miniTreeNode23);
            //孫ノード24

            miniTreeNode24.Text = "依頼文";
            treeNode3.Nodes.Add(miniTreeNode24);
            //孫ノード25

            miniTreeNode25.Text = "照会文";
            treeNode3.Nodes.Add(miniTreeNode25);
            //孫ノード26

            miniTreeNode26.Text = "回答文";
            treeNode3.Nodes.Add(miniTreeNode26);
            //孫ノード27

            miniTreeNode27.Text = "通知文";
            treeNode3.Nodes.Add(miniTreeNode27);
            //孫ノード28

            miniTreeNode28.Text = "案内文";
            treeNode3.Nodes.Add(miniTreeNode28);
            //ノード4

            treeNode4.Text = "報告文書";
            treeView2.Nodes.Add(treeNode4);
            //孫ノード29

            miniTreeNode29.Text = "参加報告書";
            treeNode4.Nodes.Add(miniTreeNode29);
            //孫ノード30

            miniTreeNode30.Text = "出張報告書";
            treeNode4.Nodes.Add(miniTreeNode30);
            //孫ノード31

            miniTreeNode31.Text = "上申書";
            treeNode4.Nodes.Add(miniTreeNode31);
            //孫ノード32

            miniTreeNode32.Text = "届出文";
            treeNode4.Nodes.Add(miniTreeNode32);
            //孫ノード33

            miniTreeNode33.Text = "始末書";
            treeNode4.Nodes.Add(miniTreeNode33);
            //孫ノード33

            miniTreeNode34.Text = "理由書";
            treeNode4.Nodes.Add(miniTreeNode34);
            //後にTreeViewをすべて展開
            treeView1.ExpandAll();
            treeView2.ExpandAll();

            kryptonNavigator1.NavigatorMode = NavigatorMode.HeaderGroup;

        }

        #endregion

        #region シートの設定
        public void SetSheets()
        {
            //置換コントロールを消す
            kryptonPanel21.Height = 0;
            //カレンダーコントロールを今日に設定する
            kryptonDateTimePicker1.CalendarTodayDate = DateTime.Now;

            //シート内の設定
            kryptonTextBox7.Text = string.Empty;
            kryptonComboBox8.Text = string.Empty;

            kryptonComboBox6.Text = string.Empty;
            kryptonTextBox14.Text = string.Empty;
            kryptonTextBox8.Text = string.Empty;

            kryptonComboBox7.Text = string.Empty;
            kryptonTextBox9.Text = string.Empty;
            kryptonTextBox15.Text = string.Empty;

            kryptonComboBox2.Text = "拝啓";

            DateTime date = DateTime.Now;
            kryptonComboBox1.Text = date.Month.ToString();

            kryptonComboBox3.Text = "貴社ますますご盛栄のこととお慶び申し上げます。";


            kryptonComboBox4.Text = "平素は格別のご高配を賜り、厚く御礼申し上げます。";


            Sheets_TitleButton.Dock = DockStyle.Fill;

            InstalledFontCollection fonts = new InstalledFontCollection();
            FontFamily[] fontFamilies = fonts.Families;


            foreach (FontFamily font in fontFamilies)
            {
                kryptonRibbonGroupComboBox_Font.Items.Add(font.Name);
                kryptonRibbonGroupComboBox_Font.AutoCompleteCustomSource.Add(font.Name);
                kryptonRibbonGroupComboBox_NotepadFont.Items.Add(font.Name);
                kryptonRibbonGroupComboBox_NotepadFont.AutoCompleteCustomSource.Add(font.Name);

                toolStripComboBox1.Items.Add(font.Name);
                toolStripComboBox2.AutoCompleteCustomSource.Add(font.Name);

                toolStripComboBox3.Items.Add(font.Name);
                toolStripComboBox4.AutoCompleteCustomSource.Add(font.Name);

                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }

            //シート内の設定を変更
            kryptonCheckBox1.Checked = true;
            kryptonCheckButton2.Checked = false;

            Sheets_NumberLabel.Visible = false;
            Sheets_DateLabel.Visible = false;
            Sheets_AddressCompanyLabel.Visible = false;
            Sheets_AddressTitleAndNameLabel.Visible = false;
            Sheets_CallerCompanyLabel.Visible = false;
            Sheets_CallerLocationLabel.Visible = false;
            Sheets_BuildingNameLabel.Visible = false;
            Sheets_CallerTitleAndNameLabel.Visible = false;
            Sheets_CallerMallAddressLabel.Visible = false;
            Sheets_CallerTelLabel.Visible = false;
            Sheets_CallerFaxTelLabel.Visible = false;
            Sheets_TitleButton.Visible = false;
            Sheets_ContentLabel.Visible = false;
            Sheet_ConclusionLabel.Visible = false;

            panel4.Height = 221;

            panel2.Visible = true;
            panel3.Visible = true;
            kryptonTextBox1.Visible = true;
            panel11.Visible = true;
            kryptonTextBox3.Visible = true;
            kryptonTextBox4.Visible = true;
            panel6.Visible = true;
            panel10.Visible = true;
            panel9.Visible = true;
            kryptonTextBox8.Visible = true;
            kryptonTextBox9.Visible = true;
            kryptonTextBox10.Visible = true;
            panel5.Visible = true;
            panel7.Visible = true;
            panel8.Visible = true;

            //編集用シートの初期化
            //日付を今日に設定する
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime date1 = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date1 = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
            }

        }
        #endregion

        public void SetDialog()
        {
            if (Properties.Settings.Default.ShowNotepadWarningPanel == true)
            {
                WarningPanel1.Visible = true;
            }
            else
            {
                WarningPanel1.Visible = false;
            }

        }

        public void IsSoftWareUpdateAvailable()
        {
            if (System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
            {
                WebClient wc = new WebClient();
                System.Version currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                System.Version updateVersion = new System.Version(wc.DownloadString("https://raw.githubusercontent.com/User233389/DocuEase/refs/heads/main/UpdateVersion.txt"));
                if (updateVersion > currentVersion)
                {
                    MessageBox.Show("新しいバージョンが公開されています。最新バージョンをダウンロードしてください。");
                }
                wc.Dispose();

            }


        }

        public void SetRibbonOrMenuBar()
        {


            //リボンかメニューバーかの設定
            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                menuStripPanel.Visible = false;
                kryptonRibbon.Visible = true;

                this.AllowFormChrome = true;

                kryptonContextMenuItem96.Checked = false;
            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                menuStripPanel.Visible = true;
                kryptonRibbon.Visible = false;

                this.AllowFormChrome = false;

                kryptonContextMenuItem96.Checked = true;
            }
        }

        public void SetRecentDoc()
        {
            kryptonRibbonRecentDoc1.Text = Properties.Settings.Default.RecentDoc1;
            kryptonRibbonRecentDoc2.Text = Properties.Settings.Default.RecentDoc2;
            kryptonRibbonRecentDoc3.Text = Properties.Settings.Default.RecentDoc3;
            kryptonRibbonRecentDoc4.Text = Properties.Settings.Default.RecentDoc4;
            kryptonRibbonRecentDoc5.Text = Properties.Settings.Default.RecentDoc5;
            kryptonRibbonRecentDoc6.Text = Properties.Settings.Default.RecentDoc6;
            kryptonRibbonRecentDoc7.Text = Properties.Settings.Default.RecentDoc7;
            kryptonRibbonRecentDoc8.Text = Properties.Settings.Default.RecentDoc8;
            kryptonRibbonRecentDoc9.Text = Properties.Settings.Default.RecentDoc9;
        }

        //アプリの読み込み処理
        private void Form1_Load(object sender, EventArgs e)
        {
            kryptonRibbon.SelectedContext = "Notepad";
            Form1.Form1Instance = this;

            //ソフトウェアの更新確認
            //IsSoftWareUpdateAvailable();

            //最近使用したテンプレートの復元
            SetRecentDoc();

            //ノードの追加処理
            AddTreeNodes();
            //シートの設定処理
            SetSheets();
            //テーマの復元
            SetTheme();
            //QATの表示状態の復元
            SetQAT();
            //シートの空白間隔の復元
            SetSheetSpace();
            //シートのテキストの復元
            SetSheetText();
            //設定画面の復元
            SetSettings();
            //DCMKの起動タスク
            RunAppTask();
            //ダイアログ表示の復元    
            SetDialog();
            //リボンかメニューバーか
            SetRibbonOrMenuBar();

            //リボンとメニューバーの切り替えに関するプロパティの初期化
            Sheets_SelectForeColor = Sheets_TitleButton.ForeColor;

            //キーボードショートカットの初期化
            kryptonRibbon.SelectedContext = string.Empty;
            kryptonRibbonButton_Paste.ShortcutKeys = Keys.Control | Keys.V;
            kryptonRibbonGroupButton_NotepadPaste.ShortcutKeys = Keys.None;

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            //メニューバーの表示制御
            コンタクトToolStripMenuItem.Visible = false;
            メモ帳ToolStripMenuItem.Visible = false;

            toolStripComboBox1.Text = Sheets_TitleButton.Font.Name.ToString();
            toolStripComboBox2.Text = Sheets_TitleButton.Font.Size.ToString();

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            //シート
            kryptonRibbonButton_Bold.ShortcutKeys = Keys.Control | Keys.B;
            kryptonRibbonButton_Italic.ShortcutKeys = Keys.Control | Keys.I;
            kryptonContextMenuItem15.ShortcutKeys = Keys.Control | Keys.U;
            kryptonContextMenuItem16.ShortcutKeys = Keys.Control | Keys.T;

            kryptonRibbonColorButton_TextColor.ShortcutKeys = Keys.Control | Keys.Shift | Keys.C;

            kryptonRibbonGroupClusterButton4.ShortcutKeys = Keys.Control | Keys.Shift | Keys.U;
            kryptonRibbonGroupClusterButton5.ShortcutKeys = Keys.Control | Keys.Shift | Keys.D;

            //メモ
            kryptonRibbonGroupClusterButton1.ShortcutKeys = Keys.None;
            kryptonRibbonGroupClusterButton2.ShortcutKeys = Keys.None;
            kryptonContextMenuItem35.ShortcutKeys = Keys.None;
            kryptonContextMenuItem36.ShortcutKeys = Keys.None;

            kryptonRibbonGroupColorButton2.ShortcutKeys = Keys.None;
            kryptonRibbonGroupColorButton3.ShortcutKeys = Keys.None;

            kryptonRibbonGroupClusterButton6.ShortcutKeys = Keys.None;
            kryptonRibbonGroupClusterButton7.ShortcutKeys = Keys.None;

            try
            {
                //メモの内容を復元
                String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase\SaveFile.rtf";

                if (File.Exists(str))
                {
                    Notepads_kryptonRichTextBox_Notepad.LoadFile(str);
                }
                else
                {
                    Notepads_kryptonRichTextBox_Notepad.Text = "(ここにメモしたい文字を入力します)";
                }
            }
            catch
            { }

        }

        private void kryptonCommandLinkButton1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonSplitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void kryptonSplitContainer1_SplitterMoving(object sender, SplitterCancelEventArgs e)
        {

        }


        private void DatePage_Click(object sender, EventArgs e)
        {

        }

        private void WarningPanel_CloseButton_Click(object sender, EventArgs e)
        {
            Transition
                .With(WarningPanel1, nameof(Height), 0)
                .CriticalDamp(TimeSpan.FromSeconds(0.4));
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
        }

        private void kryptonLabel24_Click(object sender, EventArgs e)
        {

        }



        private void kryptonNavigator_Workbench_Selecting(object sender, ComponentFactory.Krypton.Navigator.KryptonPageCancelEventArgs e)
        {

        }

        private void kryptonRibbon_SelectedTabChanged(object sender, EventArgs e)
        {

        }

        private void kryptonNavigator_Workbench_SelectedPageChanged(object sender, EventArgs e)
        {

        }




        private void kryptonRibbonButton_Content_Click(object sender, EventArgs e)
        {

        }

        #region テーマの切り替え処理
        private void of2007_Click(object sender, EventArgs e)
        {
            //青
            if (kryptonContextMenuRadioButton1.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(191, 219, 255);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(191, 219, 255);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007青色
                this.BackColor = Color.FromArgb(191, 219, 255);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(191, 219, 255);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //シルバー
            else if (kryptonContextMenuRadioButton2.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(208, 212, 221);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(208, 212, 221);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007銀色
                this.BackColor = Color.FromArgb(208, 212, 221);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(208, 212, 221);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //ブラック
            else if (kryptonContextMenuRadioButton3.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(83, 83, 83);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(83, 83, 83);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007黒色
                this.BackColor = Color.FromArgb(83, 83, 83);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(83, 83, 83);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

            }
            of2007.Checked = true;
            of2010.Checked = false;
        }

        private void of2010_Click(object sender, EventArgs e)
        {
            //青
            if (kryptonContextMenuRadioButton1.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(187, 206, 230);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(187, 206, 230);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010青色
                this.BackColor = Color.FromArgb(187, 206, 230);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(187, 206, 230);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //シルバー
            else if (kryptonContextMenuRadioButton2.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(227, 230, 232);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(227, 230, 232);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010銀色
                this.BackColor = Color.FromArgb(227, 230, 232);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(227, 230, 232);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //ブラック
            else if (kryptonContextMenuRadioButton3.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(113, 113, 113);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(113, 113, 113);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010黒色
                this.BackColor = Color.FromArgb(113, 113, 113);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(113, 113, 113);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

            }


            of2007.Checked = false;
            of2010.Checked = true;
        }
        #endregion

        #region テーマカラーの切り替えとそれに伴う処理
        private void kryptonContextMenuRadioButton1_Click(object sender, EventArgs e)
        {
            //青
            if (of2007.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(191, 219, 255);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(191, 219, 255);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007青色
                this.BackColor = Color.FromArgb(191, 219, 255);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(191, 219, 255);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //青
            else if (of2010.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Blue";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(187, 206, 230);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(187, 206, 230);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010青色
                this.BackColor = Color.FromArgb(187, 206, 230);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(187, 206, 230);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;



                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
        }

        private void kryptonContextMenuRadioButton2_Click(object sender, EventArgs e)
        {
            //シルバー
            if (of2007.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(208, 212, 221);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(208, 212, 221);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007銀色
                this.BackColor = Color.FromArgb(208, 212, 221);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(208, 212, 221);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Blue;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (of2010.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Silver";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(227, 230, 232);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(227, 230, 232);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010銀色
                this.BackColor = Color.FromArgb(227, 230, 232);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.White;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.Black;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(227, 230, 232);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;


                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
        }

        private void kryptonContextMenuRadioButton3_Click(object sender, EventArgs e)
        {
            //ブラック
            if (of2007.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2007Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(83, 83, 83);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(83, 83, 83);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.PanelClient;

                //FormBackColor
                //Word2007黒色
                this.BackColor = Color.FromArgb(83, 83, 83);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(83, 83, 83);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                WarningPanel_CloseButton.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;


                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //Ribbon
                kryptonRibbon.StateCommon.RibbonTab.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupButtonText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupNormalTitle.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCollapsedText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupCheckBoxText.TextColor = Color.Empty;
                kryptonRibbon.StateCommon.RibbonGroupLabelText.TextColor = Color.Empty;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007BlackDarkMode;


                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2007Black;

            }
            //ブラック
            else if (of2010.Checked == true)
            {
                Properties.Settings.Default.Theme = "Office2010Black";
                Properties.Settings.Default.Save();
                //SplitContainer
                kryptonSplitContainer2.StateCommon.Back.Color1 = Color.FromArgb(113, 113, 113);
                //コンテキストメニュー
                kryptonContextMenu8.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                //Command Link
                kryptonCommandLinkButton1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                kryptonCommandLinkButton2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Silver;
                //Sheet
                kryptonPage1.StateCommon.Page.Color1 = Color.FromArgb(113, 113, 113);
                kryptonPanel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonPanel2.PanelBackStyle = Krypton.Toolkit.PaletteBackStyle.HeaderSecondary;

                //FormBackColor
                //Word2010黒色
                this.BackColor = Color.FromArgb(113, 113, 113);

                //メモ
                Notepads_kryptonRichTextBox_Notepad.BackColor = Color.Black;
                Notepads_kryptonRichTextBox_Notepad.ForeColor = Color.White;

                //WarningPanel
                WarningPanel1.BackColor = Color.FromArgb(113, 113, 113);
                WarningPanel_Text.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonLinkLabel1.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //EditPanel
                kryptonPanel1.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;

                //連絡先タブ
                Address_NameLabel.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;
                kryptonSplitContainer2.Panel2.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;
                kryptonListBox3.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010BlackDarkMode;

                kryptonButton6.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Black;

            }
        }
        #endregion


        private void kryptonContextMenu1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void RibbonAppButtonContextMenu_AboutApp_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox();

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                about.BackColor = Color.FromArgb(191, 219, 255);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                about.BackColor = Color.FromArgb(208, 212, 221);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                about.BackColor = Color.FromArgb(83, 83, 83);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                about.BackColor = Color.FromArgb(187, 206, 230);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                about.BackColor = Color.FromArgb(227, 230, 232);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                about.BackColor = Color.FromArgb(113, 113, 113);
                about.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            about.ShowDialog();
        }

        private void buttonSpecAppMenu1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void kryptonContextMenuItem13_Click(object sender, EventArgs e)
        {
            ThirdParty thirdParty = new ThirdParty();

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                thirdParty.BackColor = Color.FromArgb(191, 219, 255);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                thirdParty.BackColor = Color.FromArgb(208, 212, 221);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                thirdParty.BackColor = Color.FromArgb(83, 83, 83);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                thirdParty.BackColor = Color.FromArgb(187, 206, 230);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                thirdParty.BackColor = Color.FromArgb(227, 230, 232);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                thirdParty.BackColor = Color.FromArgb(113, 113, 113);
                thirdParty.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            thirdParty.ShowDialog();
        }





        private void fontGroup_DialogBoxLauncherClick(object sender, EventArgs e)
        {
            KryptonFontDialog fd = new KryptonFontDialog();
            fd.DisplayExtendedColorsButton = true;
            fd.Font = Sheets_TitleButton.Font;
            fd.ShowColor = true;
            fd.Color = Sheets_TitleButton.ForeColor;
            kryptonRibbonColorButton_TextColor.SelectedColor = Sheets_TitleButton.ForeColor;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                Sheets_TitleButton.Font = fd.Font;
                Sheets_TitleButton.ForeColor = fd.Color;
                kryptonRibbonColorButton_TextColor.SelectedColor = fd.Color;

                kryptonTextBox10.StateCommon.Content.Font = fd.Font;
                kryptonTextBox10.StateCommon.Content.Color1 = fd.Color;

                kryptonRibbonGroupComboBox_Font.Text = fd.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = fd.Font.Size.ToString();

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Bold)
                {
                    kryptonRibbonButton_Bold.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Bold.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Italic)
                {
                    kryptonRibbonButton_Italic.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Italic.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Underline)
                {
                    kryptonContextMenuItem15.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem15.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Strikeout)
                {
                    kryptonContextMenuItem16.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem16.Checked = false;
                }
            }
        }

        public void fd_ShowHelpReqest(Object sender, EventArgs e)
        {

        }

        public void fd_ShowHelpReqest2(Object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupButton_NotepadFonts_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonColorButton_TextColor_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            Sheets_SelectForeColor = e.Color;
            Sheets_TitleButton.ForeColor = kryptonRibbonColorButton_TextColor.SelectedColor;
            kryptonTextBox10.StateCommon.Content.Color1 = kryptonRibbonColorButton_TextColor.SelectedColor;
        }


        private void Notepads_kryptonRichTextBox_Notepad_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void Notepads_kryptonRichTextBox_Notepad_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_Font_TextUpdate(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void Sheets_TitleLabel_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_Font_SelectionChangeCommitted(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_Font_DropDownClosed(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_Font_SelectedValueChanged(object sender, EventArgs e)
        {
        }

        private void kryptonContextMenuItem1_CheckStateChanged(object sender, EventArgs e)
        {

        }


        private void kryptonContextMenuItem1_Click(object sender, EventArgs e)
        {

            DCW dCW = new DCW();

            Properties.Settings.Default.dCW_TopSpace = Sheets_TopPanel.Height;
            Properties.Settings.Default.dCW_ButtomSpace = Sheets_ButtomPanel.Height;
            Properties.Settings.Default.dCW_LeftSpace = Sheets_LeftPanel.Width;
            Properties.Settings.Default.dCW_RightSpace = Sheets_RightPanel.Width;
            Properties.Settings.Default.Save();


            // Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                dCW.BackColor = Color.FromArgb(191, 219, 255);
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                dCW.BackColor = Color.FromArgb(208, 212, 221);
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                dCW.BackColor = Color.FromArgb(83, 83, 83);
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                dCW.BackColor = Color.FromArgb(187, 206, 230);
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                dCW.BackColor = Color.FromArgb(227, 230, 232);
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                dCW.BackColor = Color.FromArgb(113, 113, 113);
            }

            dCW.ShowDialog();

            //trueの場合のみ実行
            //ウィザードの「完了」ボタンをクリックしたときに実行
            if (dCW.IsWizardFinished == true)
            {
                FontReset();
                //発行番号
                if (dCW.NoIssueNumber == false)
                {
                    kryptonCheckBox3.Checked = false;
                    kryptonTextBox11.Text = dCW.IssueNumber_Publisher;
                    kryptonNumericUpDown1.Value = dCW.IssueNumber;
                }
                else
                {
                    kryptonCheckBox3.Checked = true;
                }

                //日付
                if (dCW.NoDate == false)
                {
                    kryptonCheckBox2.Checked = false;
                    kryptonDateTimePicker1.Value = dCW.Date;
                    if (dCW.UseEraName == true)
                    {
                        kryptonCheckBox1.Checked = true;
                    }
                    else
                    {
                        kryptonCheckBox1.Checked = false;
                    }
                }
                else
                {
                    kryptonCheckBox2.Checked = true;
                }

                //発信者
                kryptonTextBox1.Text = dCW.AdCompany;
                kryptonComboBox10.Text = dCW.AdTitle;
                kryptonTextBox2.Text = dCW.AdName;

                kryptonTextBox3.Text = dCW.CaCampany;
                kryptonTextBox4.Text = dCW.CaLocation;
                kryptonTextBox5.Text = dCW.CaBuildingName;
                kryptonNumericUpDown2.Value = dCW.CaFloorNumber;
                kryptonComboBox9.Text = dCW.CaTitle;
                kryptonTextBox6.Text = dCW.CaName;
                kryptonTextBox7.Text = dCW.CaMailAddress;
                kryptonComboBox8.Text = dCW.CaMailAddress_Domain;
                //電話番号
                kryptonComboBox6.Text = dCW.CaPhoneNumber1;
                kryptonTextBox14.Text = dCW.CaPhoneNumber2;
                kryptonTextBox8.Text = dCW.CaPhoneNumber3;
                kryptonComboBox7.Text = dCW.CaFaxNumber1;
                kryptonTextBox9.Text = dCW.CaFaxNumber2;
                kryptonTextBox15.Text = dCW.CaFaxNumber3;

                //表題
                kryptonTextBox10.Text = dCW.title;
                kryptonTextBox10.StateCommon.Content.Color1 = dCW.titleColor;
                Sheets_TitleButton.ForeColor = dCW.titleColor;

                //表題のフォント
                kryptonRibbonGroupComboBox_Font.Text = dCW.ftName;
                kryptonRibbonGroupComboBox_FontSize.Text = dCW.ftSize.ToString();
                kryptonRibbonColorButton_TextColor.SelectedColor = dCW.titleColor;


                if (dCW.titleBold == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Bold);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;

                    kryptonRibbonButton_Bold.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Bold.Checked = false;
                }

                if (dCW.titleItalic == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Italic);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Italic.Checked = true;
                }
                else
                {
                    kryptonRibbonButton_Italic.Checked = false;
                }

                if (dCW.titleUnderline == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Underline);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem15.Checked = true;
                }
                if (dCW.titleUnderline == false)
                {
                    kryptonContextMenuItem15.Checked = false;
                }

                if (dCW.titleStrikeout == true)
                {
                    Sheets_TitleButton.Font = new System.Drawing.Font(dCW.ftName, dCW.ftSize, Sheets_TitleButton.Font.Style | FontStyle.Strikeout);
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem16.Checked = true;
                }
                if (dCW.titleStrikeout == false)
                {
                    kryptonContextMenuItem16.Checked = false;
                }

                //あいさつ文
                //月
                kryptonComboBox1.Text = dCW.UseSourouBunDate;
                //頭語
                kryptonComboBox2.Text = dCW.acronym;
                //候文
                kryptonComboBox11.Text = dCW.souroubun;
                //前文
                kryptonComboBox3.Text = dCW.PreviousText;
                //感謝のあいさつ
                kryptonComboBox4.Text = dCW.ThankYouGreeting;
                //結語
                kryptonComboBox5.Text = dCW.Conclusion;

                //内容
                kryptonTextBox12.Text = dCW.Content;
                kryptonTextBox13.Text = dCW.Notetaking;
            }
            // falseの場合は何もしない
        }

        private void kryptonRibbonGroupComboBox_Font_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_Font_KeyDown(object sender, KeyEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    kryptonRibbonGroupComboBox_Font.Text,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupButton1_Click(object sender, EventArgs e)
        {

        }

        private void Editpanel_NoEditCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void kryptonContextMenuItem12_Click(object sender, EventArgs e)
        {

        }



        UseSoftWareWindow useSoftWareWindow = new UseSoftWareWindow();
        KeboradShortCut keboradShortCut = new KeboradShortCut();
        SheetsScaleDialog sheetsScaleDialog = new SheetsScaleDialog();
        public void Form1_Activated(object sender, EventArgs e)
        {

            if (useSoftWareWindow.Visible == true)
            {
                kryptonRibbonGroupButton_Tutorial.Enabled = false;
                kryptonContextMenuItem12.Enabled = false;
            }
            else
            {
                kryptonRibbonGroupButton_Tutorial.Enabled = true;
                kryptonContextMenuItem12.Enabled = true;
            }

            if (keboradShortCut.Visible == true)
            {
                buttonSpecAppMenu2.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.False;
                キーボードショートカットの確認ToolStripMenuItem.Enabled = false;
            }
            else
            {
                buttonSpecAppMenu2.Enabled = ComponentFactory.Krypton.Toolkit.ButtonEnabled.True;
                キーボードショートカットの確認ToolStripMenuItem.Enabled = true;
            }


            if (sheetsScaleDialog.Visible == true)
            {
                シートのサイズToolStripMenuItem.Enabled = false;
                kryptonRibbonGroup6.DialogBoxLauncher = false;
            }
            else
            {
                シートのサイズToolStripMenuItem.Enabled = true;
                kryptonRibbonGroup6.DialogBoxLauncher = true;
            }
        }
        //再起動かどうかのフラグ
        public bool IsAppRestarting { get; set; } = false;//基本的にfalseに設定しておく
        //保存ファイルを削除したか確かめるフラグ
        public bool SaveFileDeleted { get; set; } = false;
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Visible = false;
            if (SaveFileDeleted == false)
            {
                //ファイル保存
                AutoSave();
            }


            //アプリケーション再起動中でない場合
            if (IsAppRestarting == false)
            {
                //不要なオブジェクトを破棄してから終了する
                useSoftWareWindow.Dispose();
                keboradShortCut.Dispose();
                sheetsScaleDialog.Dispose();

                //QATの位置を保存
                if (kryptonRibbon.QATLocation == QATLocation.Above)
                {
                    Properties.Settings.Default.ShowQATLocation = 0;
                }
                else if (kryptonRibbon.QATLocation == QATLocation.Below)
                {
                    Properties.Settings.Default.ShowQATLocation = 1;
                }

                //QAT状態確認・保存
                if (kryptonRibbonQATButton1.Visible == true)
                {
                    Properties.Settings.Default.QAT1_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT1_Visible = false;
                }

                if (kryptonRibbonQATButton2.Visible == true)
                {
                    Properties.Settings.Default.QAT2_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT2_Visible = false;
                }

                if (kryptonRibbonQATButton3.Visible == true)
                {
                    Properties.Settings.Default.QAT3_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT3_Visible = false;
                }

                if (kryptonRibbonQATButton4.Visible == true)
                {
                    Properties.Settings.Default.QAT4_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT4_Visible = false;
                }

                if (kryptonRibbonQATButton5.Visible == true)
                {
                    Properties.Settings.Default.QAT5_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT5_Visible = false;
                }

                if (kryptonRibbonQATButton6.Visible == true)
                {
                    Properties.Settings.Default.QAT6_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT6_Visible = false;
                }

                if (kryptonRibbonQATButton7.Visible == true)
                {
                    Properties.Settings.Default.QAT7_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT7_Visible = false;
                }

                if (kryptonRibbonQATButton8.Visible == true)
                {
                    Properties.Settings.Default.QAT8_Visible = true;
                }
                else
                {
                    Properties.Settings.Default.QAT8_Visible = false;
                }

                Properties.Settings.Default.Save();

                //イベントをすべて解除
                //ClearEvant();
                //オブジェクトをすべて破棄
                this.Dispose();
                //ガベージコレクション
                GC.Collect();
            }

        }

        #region リボンコントロールのホーム「文書作成ソフトウェアで編集」をクリックしたときの処理

        private void SetWordRangeColor(Range range, Color color)
        {
            // Word の RGB 値は Red + (Green << 8) + (Blue << 16)
            int rgb = color.R | (color.G << 8) | (color.B << 16);
            range.Font.Color = (WdColor)rgb;
        }


        private void kryptonRibbonButton_OpenWSoft_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }


            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Name = Sheets_TitleButton.Font.Name;
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Name = "游明朝";
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }

            GC.Collect();





        }
        #endregion

        #region リボンコントロールのホーム「Docx  形式で保存」をクリックしたときの処理
        private void kryptonRibbonGroupButton10_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //バックグラウンド上でWordを起動する
            word.Visible = false;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }

            //保存処理
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "ドキュメントファイルを保存する場所を選択";
            sd.Filter = "Word 文書 (*.docx)|*.docx";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.SaveAs2(sd.FileName);
                    MessageBox.Show("ファイルが以下の場所に正しく保存されました。\r\n" + sd.FileName, "ファイル保存完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("ファイルが正しく保存されませんでした。保存するファイルの場所が適切か文書作成ソフトウェアがインストールされているか確認してください。\r\n\r\nエラー内容:\r\n" + ex.Message, "ファイル保存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            //保存を確認せず閉じる
            try
            {
                doc.Close(false);
                word.Quit();
            }
            catch { }


            GC.Collect();
        }
        #endregion

        //連携完了後の処理
        async System.Threading.Tasks.Task stausUpdate()
        {
            await System.Threading.Tasks.Task.Delay(5000);
            kryptonLabel1.Text = "準備完了";
        }

        public void Timer_Tick(object sender, EventArgs e)
        {




        }

        #region 表示モード切り替え処理
        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton_ViewMode1.Checked = true;
            kryptonRibbonGroupButton_ViewMode2.Checked = false;

            kryptonCheckButton1.Checked = true;
            kryptonCheckButton2.Checked = false;

            Sheets_NumberLabel.Visible = false;
            Sheets_DateLabel.Visible = false;
            Sheets_AddressCompanyLabel.Visible = false;
            Sheets_AddressTitleAndNameLabel.Visible = false;
            Sheets_CallerCompanyLabel.Visible = false;
            Sheets_CallerLocationLabel.Visible = false;
            Sheets_BuildingNameLabel.Visible = false;
            Sheets_CallerTitleAndNameLabel.Visible = false;
            Sheets_CallerMallAddressLabel.Visible = false;
            Sheets_CallerTelLabel.Visible = false;
            Sheets_CallerFaxTelLabel.Visible = false;
            Sheets_TitleButton.Visible = false;
            Sheets_ContentLabel.Visible = false;
            Sheet_ConclusionLabel.Visible = false;

            panel4.Height = 221;

            panel2.Visible = true;
            panel3.Visible = true;
            kryptonTextBox1.Visible = true;
            panel11.Visible = true;
            kryptonTextBox3.Visible = true;
            kryptonTextBox4.Visible = true;
            panel6.Visible = true;
            panel10.Visible = true;
            panel9.Visible = true;
            kryptonTextBox8.Visible = true;
            kryptonTextBox9.Visible = true;
            kryptonTextBox10.Visible = true;
            panel5.Visible = true;
            panel7.Visible = true;
            panel8.Visible = true;
            label9.Visible = true;

            kryptonComboBox5.Visible = true;

            label11.Visible = true;
            label12.Visible = true;

            Sheets_NumberPanel.Visible = true;
            Sheets_DatePanel.Visible = true;
            Sheets_AddressCompanyPanel.Visible = true;
            Sheets_AddressTitleAndNamePanel.Visible = true;
            Sheets_CallerCompanyPanel.Visible = true;
            Sheets_CallerLocationPanel.Visible = true;
            Sheets_BuildingNamePanel.Visible = true;
            Sheets_CallerTitleAndNamePanel.Visible = true;
            Sheets_CallerMallAddressPanel.Visible = true;
            Sheets_CallerTelPanel.Visible = true;
            Sheets_CallerFaxTelPanel.Visible = true;
        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {
            kryptonRibbonGroupButton_ViewMode1.Checked = false;
            kryptonRibbonGroupButton_ViewMode2.Checked = true;

            kryptonCheckButton1.Checked = false;
            kryptonCheckButton2.Checked = true;

            Sheets_NumberLabel.Visible = true;
            Sheets_DateLabel.Visible = true;
            Sheets_AddressCompanyLabel.Visible = true;
            Sheets_AddressTitleAndNameLabel.Visible = true;
            Sheets_CallerCompanyLabel.Visible = true;
            Sheets_CallerLocationLabel.Visible = true;
            Sheets_BuildingNameLabel.Visible = true;
            Sheets_CallerTitleAndNameLabel.Visible = true;
            Sheets_CallerMallAddressLabel.Visible = true;
            Sheets_CallerTelLabel.Visible = true;
            Sheets_CallerFaxTelLabel.Visible = true;
            Sheets_TitleButton.Visible = true;
            Sheets_ContentLabel.Visible = true;
            Sheet_ConclusionLabel.Visible = true;

            panel4.Height = 221;

            panel2.Visible = false;
            panel3.Visible = false;
            kryptonTextBox1.Visible = false;
            panel11.Visible = false;
            kryptonTextBox3.Visible = false;
            kryptonTextBox4.Visible = false;
            panel6.Visible = false;
            panel10.Visible = false;
            panel9.Visible = false;
            kryptonTextBox8.Visible = false;
            kryptonTextBox9.Visible = false;
            kryptonTextBox10.Visible = false;
            panel5.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;

            kryptonComboBox5.Visible = false;

            //Number
            if (kryptonCheckBox3.Checked == true)
            {
                Sheets_NumberPanel.Visible = false;
            }
            else
            {
                Sheets_NumberPanel.Visible = true;
            }

            //Date
            if (kryptonCheckBox2.Checked == true)
            {
                Sheets_DatePanel.Visible = false;
            }
            else
            {
                Sheets_DatePanel.Visible = true;
            }

            //AdCompany
            if (Sheets_AddressCompanyLabel.Text == string.Empty)
            {
                Sheets_AddressCompanyPanel.Visible = false;
            }
            else
            {
                Sheets_AddressCompanyPanel.Visible = true;
            }

            //AdName
            if (Sheets_AddressTitleAndNameLabel.Text == string.Empty)
            {
                Sheets_AddressTitleAndNamePanel.Visible = false;
            }
            else
            {
                Sheets_AddressTitleAndNamePanel.Visible = true;
            }

            //Company
            if (Sheets_CallerCompanyLabel.Text == string.Empty)
            {
                Sheets_CallerCompanyPanel.Visible = false;
            }
            else
            {
                Sheets_CallerCompanyPanel.Visible = true;
            }

            //Location
            if (Sheets_CallerLocationLabel.Text == string.Empty)
            {
                Sheets_CallerLocationPanel.Visible = false;
            }
            else
            {
                Sheets_CallerLocationPanel.Visible = true;
            }

            //Buiding Name
            if (Sheets_BuildingNameLabel.Text == string.Empty)
            {
                Sheets_BuildingNamePanel.Visible = false;
            }
            else
            {
                Sheets_BuildingNamePanel.Visible = true;
            }

            //Name
            if (Sheets_CallerTitleAndNameLabel.Text == string.Empty)
            {
                Sheets_CallerTitleAndNamePanel.Visible = false;
            }
            else
            {
                Sheets_CallerTitleAndNamePanel.Visible = true;
            }

            //Mail
            if (label11.Font.Strikeout == true)
            {
                label11.Visible = false;
                label12.Visible = false;
                Sheets_CallerMallAddressPanel.Visible = false;
            }
            else
            {
                label11.Visible = true;
                label12.Visible = true;
                Sheets_CallerMallAddressPanel.Visible = true;
            }

            //Tel
            if (label9.Font.Strikeout == true)
            {
                label9.Visible = false;
                Sheets_CallerTelPanel.Visible = false;
            }
            else
            {
                label9.Visible = true;
                Sheets_CallerTelPanel.Visible = true;
            }

            if (label10.Font.Strikeout == true)
            {
                label10.Visible = true;
                Sheets_CallerFaxTelPanel.Visible = false;
            }
            else
            {
                label10.Visible = true;
                Sheets_CallerFaxTelPanel.Visible = true;
            }
        }
        #endregion

        #region キーによる編集項目切り替え処理
        private void kryptonTextBox11_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonLabel7.Text = "シート内の項目を移動するにはFunction+Endキーを押してください。Function+Homeキーを押すと前の項目に戻ります。";
                kryptonNumericUpDown1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonLabel7.Text = "シート内の項目を移動するにはFunction+Endキーを押してください。Function+Homeキーを押すと前の項目に戻ります。";
                kryptonComboBox5.Focus();
            }
        }

        private void kryptonNumericUpDown1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonDateTimePicker1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox11.Focus();
            }
        }

        private void kryptonDateTimePicker1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonNumericUpDown1.Focus();
            }
        }

        private void kryptonTextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox10.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonDateTimePicker1.Focus();
            }
        }

        private void kryptonComboBox10_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox2.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox1.Focus();
            }
        }

        private void kryptonTextBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox3.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox10.Focus();
            }
        }

        private void kryptonTextBox3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox4.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox2.Focus();
            }
        }

        private void kryptonTextBox4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox5.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox3.Focus();
            }
        }

        //修正用
        private void kryptonTextBox5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonNumericUpDown2.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox4.Focus();
            }
        }

        private void kryptonNumericUpDown2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox9.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox5.Focus();
            }
        }

        private void kryptonComboBox9_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox6.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonNumericUpDown2.Focus();
            }
        }

        private void kryptonTextBox6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox7.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox9.Focus();
            }
        }
        //完了

        private void kryptonTextBox7_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox8.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox6.Focus();
            }
        }

        private void kryptonComboBox8_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox6.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox7.Focus();
            }
        }

        private void kryptonComboBox6_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox14.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox8.Focus();
            }
        }

        private void kryptonTextBox14_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox8.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox6.Focus();
            }
        }

        private void kryptonTextBox8_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox7.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox14.Focus();
            }
        }

        private void kryptonComboBox7_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox9.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox8.Focus();
            }
        }

        private void kryptonTextBox9_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox15.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox7.Focus();
            }
        }

        private void kryptonTextBox15_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox10.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox9.Focus();
            }
        }

        private void kryptonTextBox10_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox1.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox15.Focus();
            }
        }

        private void kryptonComboBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox2.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonTextBox10.Focus();
            }
        }

        private void kryptonComboBox2_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox11.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox1.Focus();
            }
        }



        private void kryptonComboBox11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox3.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox2.Focus();
            }
        }


        private void kryptonComboBox3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox4.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox11.Focus();
            }
        }

        private void kryptonComboBox4_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonComboBox5.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox3.Focus();
            }
        }

        private void kryptonComboBox5_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.End)
            {
                kryptonTextBox11.Focus();
            }
            else if (e.KeyCode == Keys.Home)
            {
                kryptonComboBox4.Focus();
            }
        }
        #endregion

        private void kryptonTextBox5_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void kryptonNumericUpDown2_MouseUp(object sender, MouseEventArgs e)
        {

        }

        private void kryptonTextBox11_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void kryptonTextBox11_Click(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox11_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Text != string.Empty)
            {
                label2.Text = "発第";
                Sheets_NumberLabel.Text = kryptonTextBox11.Text + "発第" + kryptonNumericUpDown1.Value + "号";
            }
            else
            {
                label2.Text = "　第";
                Sheets_NumberLabel.Text = "第" + kryptonNumericUpDown1.Value + "号";
            }

        }

        private void kryptonDateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
            }

        }

        private void kryptonTextBox1_TextChanged(object sender, EventArgs e)
        {
            Sheets_AddressCompanyLabel.Text = kryptonTextBox1.Text;
        }

        private void kryptonComboBox10_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox10.Text != string.Empty)
            {

                if (kryptonTextBox2.Text != string.Empty)
                {
                    Sheets_AddressTitleAndNameLabel.Text = kryptonComboBox10.Text + "　" + kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }

            }
            else
            {
                if (kryptonTextBox2.Text != string.Empty)
                {
                    Sheets_AddressTitleAndNameLabel.Text = kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }
            }

            if (kryptonComboBox10.Text == "お客様各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "お客様各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
            else if (kryptonComboBox10.Text == "従業員各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "従業員各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
            else
            {
                kryptonTextBox1.Enabled = true;
                kryptonTextBox2.Enabled = true;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }

        }


        private void kryptonTextBox2_TextChanged(object sender, EventArgs e)
        {

            if (kryptonComboBox10.Text != string.Empty)
            {

                if (kryptonTextBox2.Text != string.Empty)
                {
                    Sheets_AddressTitleAndNameLabel.Text = kryptonComboBox10.Text + "　" + kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }
            }
            else
            {
                if (kryptonTextBox2.Text != string.Empty)
                {

                    Sheets_AddressTitleAndNameLabel.Text = kryptonTextBox2.Text + "様";
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                }
                else
                {
                    Sheets_AddressTitleAndNameLabel.Text = string.Empty;
                    label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                }
            }

            if (kryptonComboBox10.Text == "お客様各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "お客様各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
            else if (kryptonComboBox10.Text == "従業員各位")
            {
                kryptonTextBox1.Text = string.Empty;
                kryptonTextBox1.Enabled = false;
                Sheets_AddressTitleAndNameLabel.Text = "従業員各位";
                kryptonTextBox2.Text = string.Empty;
                kryptonTextBox2.Enabled = false;
                label13.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }


        }

        private void kryptonTextBox3_TextChanged(object sender, EventArgs e)
        {
            Sheets_CallerCompanyLabel.Text = kryptonTextBox3.Text;
        }

        private void kryptonTextBox4_TextChanged(object sender, EventArgs e)
        {
            Sheets_CallerLocationLabel.Text = kryptonTextBox4.Text;
        }

        private void kryptonTextBox5_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox5.Text != string.Empty)
            {
                kryptonNumericUpDown2.Enabled = true;
                Sheets_BuildingNameLabel.Text = kryptonTextBox5.Text + "　" + kryptonNumericUpDown2.Value + "階";
                label5.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }
            else
            {
                kryptonNumericUpDown2.Enabled = false;
                Sheets_BuildingNameLabel.Text = string.Empty;
                label5.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }

        }

        private void kryptonNumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox5.Text != string.Empty)
            {
                kryptonNumericUpDown2.Enabled = true;
                Sheets_BuildingNameLabel.Text = kryptonTextBox5.Text + "　" + kryptonNumericUpDown2.Value + "階";
                label5.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                if (kryptonNumericUpDown2.Value <= 0)
                {
                    int negativeNumber = (int)kryptonNumericUpDown2.Value;
                    int positiveNumber = Math.Abs(negativeNumber);

                    Sheets_BuildingNameLabel.Text = kryptonTextBox5.Text + "　" + "地下" + positiveNumber + "階";
                }
            }

            //0の値を入力しないようにする
            if (kryptonNumericUpDown2.Value == 0)
            {
                kryptonNumericUpDown2.Value = 1;

            }
        }

        private void kryptonComboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox6_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox9.Text != string.Empty)
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonComboBox9.Text + "　" + kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }
            }
            else
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }
            }
        }

        private void kryptonComboBox9_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox9.Text != string.Empty)
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonComboBox9.Text + "　" + kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }

            }
            else
            {
                Sheets_CallerTitleAndNameLabel.Text = kryptonTextBox6.Text;
                if (kryptonTextBox6.Text == string.Empty)
                {
                    Sheets_CallerTitleAndNameLabel.Text = string.Empty;
                }
            }

        }

        private void kryptonTextBox7_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox7.Text != string.Empty)
            {
                Sheets_CallerMallAddressLabel.Text = kryptonTextBox7.Text + "@" + kryptonComboBox8.Text;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }
            else
            {
                Sheets_CallerMallAddressLabel.Text = string.Empty;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
        }

        private void kryptonComboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox8_TextChanged(object sender, EventArgs e)
        {
            if (kryptonComboBox8.Text != string.Empty)
            {
                Sheets_CallerMallAddressLabel.Text = kryptonTextBox7.Text + "@" + kryptonComboBox8.Text;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            }
            else
            {
                Sheets_CallerMallAddressLabel.Text = string.Empty;
                label11.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                label12.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
            }
        }

        private void kryptonComboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonCheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox3.Checked == true)
            {
                kryptonTextBox11.Enabled = false;
                kryptonNumericUpDown1.Enabled = false;
                label1.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                label2.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_NumberLabel.Text = string.Empty;
            }
            else
            {
                kryptonTextBox11.Enabled = true;
                kryptonNumericUpDown1.Enabled = true;
                label1.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                label2.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);

                if (kryptonTextBox11.Text != string.Empty)
                {
                    label2.Text = "発第";
                    Sheets_NumberLabel.Text = kryptonTextBox11.Text + "発第" + kryptonNumericUpDown1.Value + "号";
                }
                else
                {
                    label2.Text = "　第";
                    Sheets_NumberLabel.Text = "第" + kryptonNumericUpDown1.Value + "号";
                }

            }
        }

        private void kryptonCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox2.Checked == true)
            {
                kryptonCheckBox1.Enabled = false;
                kryptonDateTimePicker1.Enabled = false;
                Sheets_DateLabel.Text = string.Empty;
            }
            else
            {
                kryptonCheckBox1.Enabled = true;
                kryptonDateTimePicker1.Enabled = true;

                if (kryptonCheckBox1.Checked == true)
                {
                    DateTime date = kryptonDateTimePicker1.Value.Date;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    //下記のように西暦ではなく和暦として表示するように設定する
                    culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                    Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
                }
                else
                {
                    DateTime date = kryptonDateTimePicker1.Value.Date;
                    CultureInfo culturejp = new CultureInfo("ja-Jp");
                    Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
                }
            }
        }

        private void kryptonComboBox6_TextChanged(object sender, EventArgs e)
        {
            //6
            if (kryptonComboBox6.Text != string.Empty)
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                //14
                if (kryptonTextBox14.Text != string.Empty)
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    //8
                    if (kryptonTextBox8.Text != string.Empty)
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    }
                    else
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerTelLabel.Text = string.Empty;
            }

        }

        private void kryptonTextBox14_TextChanged(object sender, EventArgs e)
        {
            //6
            if (kryptonComboBox6.Text != string.Empty)
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                //14
                if (kryptonTextBox14.Text != string.Empty)
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    //8
                    if (kryptonTextBox8.Text != string.Empty)
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    }
                    else
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerTelLabel.Text = string.Empty;
            }
        }

        private void kryptonTextBox8_TextChanged(object sender, EventArgs e)
        {
            //6
            if (kryptonComboBox6.Text != string.Empty)
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                //14
                if (kryptonTextBox14.Text != string.Empty)
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    //8
                    if (kryptonTextBox8.Text != string.Empty)
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerTelLabel.Text = kryptonComboBox6.Text + "-" + kryptonTextBox14.Text + "-" + kryptonTextBox8.Text;
                    }
                    else
                    {
                        label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label9.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerTelLabel.Text = string.Empty;
            }
        }


        private void kryptonComboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox7_TextChanged(object sender, EventArgs e)
        {
            //7
            if (kryptonComboBox7.Text != string.Empty)
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                //15
                if (kryptonTextBox15.Text != string.Empty)
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    //9
                    if (kryptonComboBox9.Text != string.Empty)
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    }
                    else
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerFaxTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerFaxTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerFaxTelLabel.Text = string.Empty;

            }
        }

        private void kryptonTextBox9_TextChanged(object sender, EventArgs e)
        {
            //9
            if (kryptonTextBox9.Text != string.Empty)
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                //15
                if (kryptonTextBox15.Text != string.Empty)
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    //7
                    if (kryptonComboBox7.Text != string.Empty)
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    }
                    else
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerFaxTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerFaxTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerFaxTelLabel.Text = string.Empty;
            }
        }

        private void kryptonTextBox15_TextChanged(object sender, EventArgs e)
        {
            //15
            if (kryptonTextBox15.Text != string.Empty)
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                //9
                if (kryptonTextBox9.Text != string.Empty)
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    //7
                    if (kryptonComboBox7.Text != string.Empty)
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                        Sheets_CallerFaxTelLabel.Text = kryptonComboBox7.Text + "-" + kryptonTextBox9.Text + "-" + kryptonTextBox15.Text;
                    }
                    else
                    {
                        label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                        Sheets_CallerFaxTelLabel.Text = string.Empty;
                    }
                }
                else
                {
                    label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    Sheets_CallerFaxTelLabel.Text = string.Empty;
                }
            }
            else
            {
                label10.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                Sheets_CallerFaxTelLabel.Text = string.Empty;
            }
        }

        private void Notepads_kryptonRichTextBox_Notepad_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;
            }
            else
            {
                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;
            }

            if (this.Width <= 902)
            {
                Sheets_Sheet.Top = 59;
                Sheets_Sheet.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            }
            else
            {

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

            }
        }

        private void kryptonComboBox2_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
            #region 頭語の選択による結語候補の切り替え処理
            //一般的
            if (kryptonComboBox2.Text == "拝啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "拝呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "啓上")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "敬白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "拝進")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            //丁寧さ
            else if (kryptonComboBox2.Text == "謹啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "謹呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "粛啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "慕啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            else if (kryptonComboBox2.Text == "謹白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "謹言",
                    "敬白",
                    "再拝",
                    "頓首",
                });
                kryptonComboBox5.Text = "謹言";
            }
            //急ぎ
            else if (kryptonComboBox2.Text == "急啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "不二",
                    "早々",
                    "不備",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "急呈")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "不二",
                    "早々",
                    "不備",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "急白")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "不二",
                    "早々",
                    "不備",
                });
                kryptonComboBox5.Text = "草々";
            }
            //略式
            else if (kryptonComboBox2.Text == "前略")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "冠省")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "略啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "寸啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            else if (kryptonComboBox2.Text == "草啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "草々",
                    "不一",
                    "不尽",
                    "早々",
                    "不二",
                });
                kryptonComboBox5.Text = "草々";
            }
            //初めて
            else if (kryptonComboBox2.Text == "初めてお手紙を差し上げます")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "突然お手紙を差し上げますご無礼お許しください")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝具",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            //重ねて
            else if (kryptonComboBox2.Text == "拝復")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "複啓")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            else if (kryptonComboBox2.Text == "謹復")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                    "拝答",
                    "敬白",
                });
                kryptonComboBox5.Text = "敬具";
            }
            //お悔み
            else if (kryptonComboBox2.Text == "合掌")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "合掌",
                });
                kryptonComboBox5.Text = "合掌";
            }
            else if (kryptonComboBox2.Text == "敬具")
            {
                kryptonComboBox5.Items.Clear();
                kryptonComboBox5.Items.AddRange(new object[] {
                    "敬具",
                });
                kryptonComboBox5.Text = "敬具";
            }
            #endregion
        }

        private void kryptonComboBox1_TextChanged(object sender, EventArgs e)
        {
            #region 月の選択による候文候補の切り替え処理
            if (kryptonComboBox1.Text == "1")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "新春の候、",
                    "初春の候、",
                    "頌春の候、",
                    "厳寒の候、",
                    "厳冬の候、",
                    "中冬の候、",
                    "寒冷の候、",
                    "麗春の候、",
                    "大寒のみぎり、",
                    "酷寒のみぎり、",
                    "寒さ厳しき季節、",
                });
                kryptonComboBox11.Text = "新春の候、";
            }
            else if (kryptonComboBox1.Text == "2")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "余寒の候、",
                    "春寒の候、",
                    "晩冬の候、",
                    "向春の候、",
                    "解氷の候、",
                    "梅花の候、",
                    "余寒なお厳しき折、",
                });
                kryptonComboBox11.Text = "余寒の候、";
            }
            else if (kryptonComboBox1.Text == "3")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "早春の候、",
                    "春寒の候、",
                    "孟春の候、",
                    "春雨降りやまぬ候、",
                    "浅春のみぎり、",
                    "春寒しだいに緩むころ、",
                    "冬の名残のまだ去りやらぬ時候、",
                    "春光天地に満ちて快い時候、",
                    "春分の季節、",
                    "春色のなごやかな季節、",
                });
                kryptonComboBox11.Text = "早春の候、";
            }
            else if (kryptonComboBox1.Text == "4")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "陽春の候、",
                    "春暖の候、",
                    "軽暖の候、",
                    "麗春の候、",
                    "春暖快適の候、",
                    "桜花爛漫の候、",
                    "花信相次ぐ候、",
                    "春眠暁を覚えずの候、",
                    "仲春四月、",
                    "春たけなわの今日この頃、",
                });
                kryptonComboBox11.Text = "早春の候、";
            }
            else if (kryptonComboBox1.Text == "5")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "新緑の候、",
                    "薫風の候、",
                    "初夏の候、",
                    "立夏の候、",
                    "暮春の候、",
                    "老春の候、",
                    "軽暑の候、",
                    "惜春のみぎり、",
                    "若葉の鮮やかな季節、",
                });
                kryptonComboBox11.Text = "新緑の候、";
            }
            else if (kryptonComboBox1.Text == "6")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "梅雨の候、",
                    "初夏の候、",
                    "短夜の候、",
                    "五月雨の候、",
                    "長雨の候、",
                    "薄暑の候、",
                    "向夏の候、",
                    "麦秋の候、",
                    "向暑のみぎり、",
                    "若鮎おどる季節、",
                });
                kryptonComboBox11.Text = "梅雨の候、";
            }
            else if (kryptonComboBox1.Text == "7")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "猛暑の候、",
                    "酷暑の候、",
                    "炎暑の候、",
                    "盛夏の候、",
                    "大暑の候、",
                    "灼熱の候、",
                    "炎熱のみぎり、",
                    "甚暑のみぎり、",
                    "三伏のみぎり、",
                    "暑さ厳しき折から、",
                });
                kryptonComboBox11.Text = "猛暑の候、";
            }
            else if (kryptonComboBox1.Text == "8")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "残暑の候、",
                    "残炎の候、",
                    "残夏の候、",
                    "暮夏の候、",
                    "季夏の候、",
                    "新涼の候、",
                    "秋暑厳しき候、",
                    "晩夏のみぎり、",
                    "処暑のみぎり、",
                    "処暑のみぎり、",
                });
                kryptonComboBox11.Text = "残暑の候、";
            }
            else if (kryptonComboBox1.Text == "9")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "初秋の候、",
                    "仲秋の候、",
                    "錦秋の候、",
                    "寒露の候、",
                    "黄葉の候、",
                    "秋雨の候、",
                    "金風の候、",
                    "秋晴れの候、",
                    "菊薫る候、",
                    "秋たけなわの候、",
                    "紅葉の季節、",
                    "秋冷の心地よい季節、",
                });
                kryptonComboBox11.Text = "初秋の候、";
            }
            else if (kryptonComboBox1.Text == "10")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "秋冷の候、",
                    "仲秋の候、",
                    "錦秋の候、",
                    "寒露の候、",
                    "黄葉の候、",
                    "秋雨の候、",
                    "金風の候、",
                    "秋晴れの候、",
                    "菊薫る候、",
                    "秋たけなわの候、",
                    "紅葉の季節、",
                    "秋冷の心地よい季節、",
                });
                kryptonComboBox11.Text = "初秋の候、";
            }
            else if (kryptonComboBox1.Text == "11")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "晩秋の候、",
                    "暮秋の候、",
                    "向寒の候、",
                    "深冷の候、",
                    "菊花の候、",
                    "紅葉の候、",
                    "初霜の候、",
                    "氷雨の候、",
                    "枯れ葉舞う季節、",
                });
                kryptonComboBox11.Text = "晩秋の候、";
            }
            else if (kryptonComboBox1.Text == "12")
            {
                kryptonComboBox11.Items.Clear();
                kryptonComboBox11.Items.AddRange(new object[] {
                    "寒冷の候、",
                    "師走の候、",
                    "初冬の候、",
                    "寒気の候、",
                    "霜気の候、",
                    "霜寒の候、",
                    "季冬の候、",
                    "歳晩の候、",
                    "歳末ご多忙の折、",
                    "心せわしい年の暮れ、",
                });
                kryptonComboBox11.Text = "寒冷の候、";
            }
            #endregion
        }

        #region ナビゲーションバーのサイズ切り替え処理
        private void buttonSpecNavigator1_Click(object sender, EventArgs e)
        {
            if (kryptonSplitContainer2.SplitterDistance <= 100)
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.Inherit;
                buttonSpecNavigator1.Text = "縮小";

                kryptonPage4.MaximumSize = new Size(0, 0);
                kryptonPage4.MinimumSize = new Size(0, 0);

                Transition
                    .With(kryptonSplitContainer2, nameof(kryptonSplitContainer2.SplitterDistance), 302)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }
            else
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookMini;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.FixedLeft;
                buttonSpecNavigator1.Text = "広げる";

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;
                kryptonPage4.MaximumSize = new Size(270, 0);
                kryptonPage4.MinimumSize = new Size(270, 0);

                Transition
                    .With(kryptonSplitContainer2, nameof(kryptonSplitContainer2.SplitterDistance), 42)
                    .CriticalDamp(TimeSpan.FromSeconds(0.6));
            }

        }

        private void kryptonSplitContainer2_SplitterMoving(object sender, SplitterCancelEventArgs e)
        {
            if (kryptonSplitContainer2.SplitterDistance <= 100)
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookMini;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.FixedLeft;
                buttonSpecNavigator1.Text = "広げる";

                kryptonPage4.MaximumSize = new Size(270, 0);
                kryptonPage4.MinimumSize = new Size(270, 0);

            }
            else
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;


                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.Inherit;
                buttonSpecNavigator1.Text = "縮小";

                kryptonPage4.MaximumSize = new Size(0, 0);
                kryptonPage4.MinimumSize = new Size(0, 0);
            }
        }

        private void kryptonSplitContainer2_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (kryptonSplitContainer2.SplitterDistance <= 100)
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookMini;

                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.FixedLeft;
                buttonSpecNavigator1.Text = "広げる";

                kryptonPage4.MaximumSize = new Size(270, 0);
                kryptonPage4.MinimumSize = new Size(270, 0);
            }
            else
            {
                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;

                Address_kryptonNavigator1.NavigatorMode = ComponentFactory.Krypton.Navigator.NavigatorMode.OutlookFull;


                Address_kryptonNavigator1.Button.ButtonDisplayLogic = ComponentFactory.Krypton.Navigator.ButtonDisplayLogic.None;

                buttonSpecNavigator1.Orientation = ComponentFactory.Krypton.Toolkit.PaletteButtonOrientation.Inherit;
                buttonSpecNavigator1.Text = "縮小";

                kryptonPage4.MaximumSize = new Size(0, 0);
                kryptonPage4.MinimumSize = new Size(0, 0);
            }
        }
        #endregion

        private void kryptonRibbonGroupButton7_Click(object sender, EventArgs e)
        {

        }



        private void kryptonComboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void kryptonComboBox11_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
        }

        private void kryptonComboBox3_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
        }

        private void kryptonComboBox4_TextChanged(object sender, EventArgs e)
        {
            Sheets_ContentLabel.Text = kryptonComboBox2.Text + "　" + kryptonComboBox11.Text + kryptonComboBox3.Text + kryptonComboBox4.Text;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            Sheets_Sheet.Top = 59;
            Sheets_Sheet.Left = 0;
        }

        private void kryptonComboBox5_TextChanged(object sender, EventArgs e)
        {
            Sheet_ConclusionLabel.Text = kryptonComboBox5.Text;
        }

        private void kryptonTextBox10_TextChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuEase Designer";
            }
            else
            {
                this.Text = "無題 - DocuEase Designer";
            }

        }

        private void kryptonTextBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonTextBox13_TextChanged(object sender, EventArgs e)
        {

        }



        #region クリップボード
        //コピー
        private void kryptonRibbonGroupButton1_Click_1(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Copy();
            }
            else if (kryptonNumericUpDown1.Focused == true)
            {
                //kryptonNumericUpDownにはCopyメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown1.Value.ToString();
                Clipboard.SetText(Clip);
            }
            else if (kryptonDateTimePicker1.Focused == true)
            {
                //kryptonDateTimePickerにはCopyメソッドがないためSheets_DateLabelをSetText経由でクリップボードにコピーする
                Clipboard.SetText(Sheets_DateLabel.Text);
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Copy();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox10.SelectedText);
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Copy();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Copy();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Copy();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Copy();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Copy();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //kryptonNumericUpDownにはCopyメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown2.Value.ToString();
                Clipboard.SetText(Clip);
            }
            else if (kryptonComboBox9.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox9.SelectedText);
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Copy();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Copy();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox8.SelectedText);
            }
            else if (kryptonComboBox6.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox6.SelectedText);
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Copy();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Copy();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox7.SelectedText);
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Copy();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Copy();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Copy();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox2.SelectedText);
            }
            else if (kryptonComboBox11.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox11.SelectedText);
            }
            else if (kryptonComboBox3.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox3.SelectedText);
            }
            else if (kryptonComboBox4.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox4.SelectedText);
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Copy();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox5.SelectedText);
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Copy();
            }
        }

        //切り取り
        private void kryptonRibbonGroupButton11_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Cut();
            }
            else if (kryptonNumericUpDown1.Focused == true)
            {
                //kryptonNumericUpDownにはCutメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown1.Value.ToString();
                Clipboard.SetText(Clip);
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown1.Value = 1;
            }
            else if (kryptonDateTimePicker1.Focused == true)
            {
                //kryptonDateTimePickerにはCutメソッドがないためSheets_DateLabelをSetText経由でクリップボードにコピーする
                Clipboard.SetText(Sheets_DateLabel.Text);
                //kryptonDateTimePicker1は値を削除できないためかわりに今日の日付に変更する
                kryptonDateTimePicker1.Value = kryptonDateTimePicker1.CalendarTodayDate;
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Cut();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox10.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox10.SelectedText))
                {
                    kryptonComboBox10.Text = kryptonComboBox10.Text.Replace(kryptonComboBox10.SelectedText, "");
                }
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Cut();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Cut();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Cut();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Cut();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Cut();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //kryptonNumericUpDownにはCutメソッドがないためstring＆SetText経由でクリップボードにコピーする
                string Clip = kryptonNumericUpDown2.Value.ToString();
                Clipboard.SetText(Clip);
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown2.Value = 1;
            }
            else if (kryptonComboBox9.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox9.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox9.SelectedText))
                {
                    kryptonComboBox9.Text = kryptonComboBox9.Text.Replace(kryptonComboBox9.SelectedText, "");
                }
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Cut();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Cut();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox8.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox8.SelectedText))
                {
                    kryptonComboBox8.Text = kryptonComboBox8.Text.Replace(kryptonComboBox8.SelectedText, "");
                }
            }
            else if (kryptonComboBox6.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox6.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox6.SelectedText))
                {
                    kryptonComboBox6.Text = kryptonComboBox6.Text.Replace(kryptonComboBox6.SelectedText, "");
                }
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Cut();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Cut();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox7.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox7.SelectedText))
                {
                    kryptonComboBox7.Text = kryptonComboBox7.Text.Replace(kryptonComboBox7.SelectedText, "");
                }

            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Cut();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Cut();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Cut();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox2.SelectedText);
                // 修正: '-=' 演算子はstring型に使えません。選択されたテキストを除去するにはReplaceを使います。
                if (!string.IsNullOrEmpty(kryptonComboBox2.SelectedText))
                {
                    kryptonComboBox2.Text = kryptonComboBox2.Text.Replace(kryptonComboBox2.SelectedText, "");
                }
            }
            else if (kryptonComboBox11.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox11.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox11.SelectedText))
                {
                    kryptonComboBox11.Text = kryptonComboBox11.Text.Replace(kryptonComboBox11.SelectedText, "");
                }
            }
            else if (kryptonComboBox3.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox3.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox3.SelectedText))
                {
                    kryptonComboBox3.Text = kryptonComboBox3.Text.Replace(kryptonComboBox3.SelectedText, "");
                }
            }
            else if (kryptonComboBox4.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox4.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox4.SelectedText))
                {
                    kryptonComboBox4.Text = kryptonComboBox4.Text.Replace(kryptonComboBox4.SelectedText, "");
                }
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Cut();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                Clipboard.SetText(kryptonComboBox5.SelectedText);
                if (!string.IsNullOrEmpty(kryptonComboBox5.SelectedText))
                {
                    kryptonComboBox5.Text = kryptonComboBox5.Text.Replace(kryptonComboBox5.SelectedText, "");
                }
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Cut();
            }
        }

        //削除
        private void kryptonRibbonGroupButton12_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Clear();
            }
            else if (kryptonNumericUpDown1.Focused == true)
            {
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown1.Value = 1;
            }
            else if (kryptonDateTimePicker1.Focused == true)
            {
                //kryptonDateTimePicker1は値を削除できないためかわりに今日の日付に変更する
                kryptonDateTimePicker1.Value = kryptonDateTimePicker1.CalendarTodayDate;
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Clear();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox10.SelectedText))
                {
                    kryptonComboBox10.Text = kryptonComboBox10.Text.Replace(kryptonComboBox10.SelectedText, "");
                }
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Clear();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Clear();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Clear();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Clear();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Clear();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //kryptonNumericUpDownは値を削除できないためかわりに値を1に変更する
                kryptonNumericUpDown2.Value = 1;
            }
            else if (kryptonComboBox9.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox8.SelectedText))
                {
                    kryptonComboBox8.Text = kryptonComboBox8.Text.Replace(kryptonComboBox9.SelectedText, "");
                }
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Clear();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Clear();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox8.SelectedText))
                {
                    kryptonComboBox8.Text = kryptonComboBox8.Text.Replace(kryptonComboBox8.SelectedText, "");
                }
            }
            else if (kryptonComboBox6.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox6.SelectedText))
                {
                    kryptonComboBox6.Text = kryptonComboBox6.Text.Replace(kryptonComboBox6.SelectedText, "");
                }
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Clear();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Clear();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox7.SelectedText))
                {
                    kryptonComboBox7.Text = kryptonComboBox7.Text.Replace(kryptonComboBox7.SelectedText, "");
                }
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Clear();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Clear();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Clear();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox2.SelectedText))
                {
                    kryptonComboBox2.Text = kryptonComboBox2.Text.Replace(kryptonComboBox2.SelectedText, "");
                }
            }
            else if (kryptonComboBox11.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox11.SelectedText))
                {
                    kryptonComboBox11.Text = kryptonComboBox11.Text.Replace(kryptonComboBox11.SelectedText, "");
                }
            }
            else if (kryptonComboBox3.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox3.SelectedText))
                {
                    kryptonComboBox3.Text = kryptonComboBox3.Text.Replace(kryptonComboBox3.SelectedText, "");
                }
            }
            else if (kryptonComboBox4.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox4.SelectedText))
                {
                    kryptonComboBox4.Text = kryptonComboBox4.Text.Replace(kryptonComboBox4.SelectedText, "");
                }
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Clear();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                if (!string.IsNullOrEmpty(kryptonComboBox5.SelectedText))
                {
                    kryptonComboBox5.Text = kryptonComboBox5.Text.Replace(kryptonComboBox5.SelectedText, "");
                }
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Clear();
            }
        }

        //貼り付け
        private void kryptonRibbonButton_Paste_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Paste();
            }
            //NnumricUpDawnも無視
            //DateTimePickarは無視
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Paste();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox10.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Paste();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Paste();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Paste();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Paste();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Paste();
            }
            else if (kryptonNumericUpDown2.Focused == true)
            {
                //最後にコピーした文字をkryptonNumericUpDownにペーストする
                kryptonNumericUpDown2.Value = Clipboard.GetText().Length;
            }
            else if (kryptonComboBox9.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox9.Text += Clipboard.GetText().Length;
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Paste();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Paste();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox8.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox6.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox6.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Paste();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Paste();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox7.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Paste();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Paste();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Paste();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox2.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox11.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox11.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox3.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox3.Text += Clipboard.GetText();
            }
            else if (kryptonComboBox4.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox3.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Paste();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                //最後にコピーした文字を入力済みの文字と一緒にペースト
                kryptonComboBox3.Text += Clipboard.GetText();
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Paste();
            }
        }
        #endregion

        private void kryptonRibbonGroupButton12_DropDown(object sender, ComponentFactory.Krypton.Toolkit.ContextMenuArgs e)
        {

        }

        #region 設定画面表示処理
        private void kryptonContextMenuItem4_Click(object sender, EventArgs e)
        {
            kryptonPanel21.Hide();

            menuStripPanel.Hide();
            kryptonRibbon.Hide();
            this.AllowFormChrome = false;

            kryptonPanel5.Hide();

            kryptonTrackBar1.Enabled = false;
            kryptonButton15.Enabled = false;
            kryptonButton14.Enabled = false;
            kryptonLabel42.Enabled = false;

            kryptonPage2.Visible = true;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.Panel;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage2;

            this.Text = "設定 - DocuEase Designer";

            kryptonCheckButton3.Checked = true;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage6;
        }



        private void kryptonGroupBox1_Panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = true;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage6;
        }

        private void kryptonCheckButton4_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = true;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage7;
        }

        private void kryptonCheckButton5_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = true;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage10;
        }

        private void kryptonCheckButton6_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = true;
            kryptonCheckButton7.Checked = false;

            kryptonNavigator1.SelectedPage = kryptonPage11;
        }

        private void kryptonCheckButton7_Click(object sender, EventArgs e)
        {
            kryptonCheckButton3.Checked = false;
            kryptonCheckButton4.Checked = false;
            kryptonCheckButton5.Checked = false;
            kryptonCheckButton6.Checked = false;
            kryptonCheckButton7.Checked = true;

            kryptonNavigator1.SelectedPage = kryptonPage12;
        }

        #endregion

        #region テンプレート選択画面表示処理
        private void kryptonRibbonRecentDoc10_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                kryptonRibbon.Enabled = false;
                kryptonRibbon.MinimizedMode = true;

            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                menuStripPanel.Enabled = false;
            }



            kryptonTrackBar1.Enabled = false;
            kryptonButton15.Enabled = false;
            kryptonButton14.Enabled = false;
            kryptonLabel42.Enabled = false;

            kryptonPage9.Visible = true;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage9;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.Panel;
            this.Text = "テンプレート - DocuEase Designer";

            kryptonLabel7.Enabled = false;
            kryptonCheckButton1.Enabled = false;
            kryptonCheckButton2.Enabled = false;
            kryptonLabel1.Enabled = false;
            kryptonPanel21.Hide();
        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            kryptonPanel21.Show();

            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                kryptonRibbon.Enabled = true;
                this.AllowFormChrome = true;
                kryptonRibbon.MinimizedMode = false;
            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                menuStripPanel.Enabled = true;
                this.AllowFormChrome = false;
            }





            if (kryptonPanel21.Height == 36)
            {
                kryptonPanel21.Height = 36;
                kryptonRibbonGroupButton16.Checked = true;
            }
            else if (kryptonPanel21.Height == 0)
            {
                kryptonPanel21.Height = 0;
                kryptonRibbonGroupButton16.Checked = false;
            }

            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }

            kryptonPage9.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;



            kryptonLabel7.Enabled = true;
            kryptonCheckButton1.Enabled = true;
            kryptonCheckButton2.Enabled = true;
            kryptonLabel1.Enabled = true;

            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuEase Designer";
            }
            else
            {
                this.Text = "無題 - DocuEase Designer";
            }

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;
        }
        #endregion

        //連絡帳表示処理
        private void kryptonRibbonGroupButton18_Click(object sender, EventArgs e)
        {
            kryptonNavigator_Workbench.SelectedPage = AddressTab;
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            kryptonPanel21.Show();

            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                menuStripPanel.Hide();
                kryptonRibbon.Show();
                this.AllowFormChrome = true;
            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                menuStripPanel.Show();
                kryptonRibbon.Hide();
                this.AllowFormChrome = false;
            }


            kryptonRibbon.Show();
            this.AllowFormChrome = true;

            kryptonPanel5.Show();

            //変更した設定の保存処理を行う
            if (kryptonRadioButton3.Checked == true)
            {
                Properties.Settings.Default.ShowApplicationTask = 0;
            }
            else if (kryptonRadioButton2.Checked == true)
            {
                Properties.Settings.Default.ShowApplicationTask = 1;
            }
            else if (kryptonRadioButton1.Checked == true)
            {
                Properties.Settings.Default.ShowApplicationTask = 2;
            }

            if (kryptonCheckBox4.Checked == true)
            {
                Properties.Settings.Default.IsAvailableDocumentCreationSoftware = true;
            }
            else
            {
                Properties.Settings.Default.IsAvailableDocumentCreationSoftware = false;
            }

            if (kryptonCheckBox7.Checked == true)
            {
                Properties.Settings.Default.IsUseEraName = true;
            }
            else
            {
                Properties.Settings.Default.IsUseEraName = false;
            }

            if (kryptonCheckBox5.Checked == true)
            {
                try
                {
                    Microsoft.Win32.RegistryKey regkey =
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\RunOnce", true);
                    regkey.SetValue(System.Windows.Forms.Application.ProductName, System.Windows.Forms.Application.ExecutablePath);
                    regkey.Close();
                }
                catch { }

                Properties.Settings.Default.IsWindowsStartUpRunForDCMK = true;
            }
            else
            {
                try
                {
                    Microsoft.Win32.RegistryKey regkey =
                        Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\RunOnce", true);
                    regkey.DeleteValue(System.Windows.Forms.Application.ProductName, false);
                    regkey.Close();
                }
                catch { }

                Properties.Settings.Default.IsWindowsStartUpRunForDCMK = false;
            }
            //シートの空白間隔
            Properties.Settings.Default.Space_Top = (int)kryptonNumericUpDown4.Value;
            Properties.Settings.Default.Space_Buttom = (int)kryptonNumericUpDown7.Value;
            Properties.Settings.Default.Space_Left = (int)kryptonNumericUpDown5.Value;
            Properties.Settings.Default.Space_Right = (int)kryptonNumericUpDown6.Value;

            //内容
            Properties.Settings.Default.SendingDepartment = kryptonTextBox16.Text;
            Properties.Settings.Default.To_CompanyOrOrganizationName = kryptonTextBox17.Text;
            Properties.Settings.Default.To_Title = kryptonComboBox12.Text;
            Properties.Settings.Default.To_Name = kryptonTextBox18.Text;
            Properties.Settings.Default.Caller_CompanyOrOrganizationName = kryptonTextBox19.Text;
            Properties.Settings.Default.Caller_Location = kryptonTextBox32.Text;
            Properties.Settings.Default.Caller_BuildingName = kryptonTextBox20.Text;
            Properties.Settings.Default.Caller_FloorNumber = (int)kryptonNumericUpDown3.Value;
            Properties.Settings.Default.Caller_Title = kryptonComboBox13.Text;
            Properties.Settings.Default.Caller_Name = kryptonTextBox21.Text;
            Properties.Settings.Default.Caller_MailAddress_User = kryptonTextBox22.Text;
            Properties.Settings.Default.Caller_MailAddress_Domain = kryptonComboBox14.Text;
            Properties.Settings.Default.Caller_PhoneNumber1 = kryptonComboBox15.Text;
            Properties.Settings.Default.Caller_PhoneNumber2 = kryptonTextBox23.Text;
            Properties.Settings.Default.Caller_PhoneNumber3 = kryptonTextBox24.Text;
            Properties.Settings.Default.Caller_FaxNumber1 = kryptonComboBox16.Text;
            Properties.Settings.Default.Caller_FaxNumber2 = kryptonTextBox26.Text;
            Properties.Settings.Default.Caller_FaxNumber3 = kryptonTextBox25.Text;

            Properties.Settings.Default.Save();
            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }

            kryptonPage2.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;

            if (kryptonRibbonGroupButton16.Checked == true)
            {
                kryptonPanel21.Height = 36;
            }
            else if (kryptonRibbonGroupButton16.Checked == false)
            {
                kryptonPanel21.Height = 0;
            }


            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuEase Designer";
            }
            else
            {
                this.Text = "無題 - DocuEase Designer";
            }

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void kryptonRibbonColorButton_TextColor_SelectedColorChanged(object sender, EventArgs e)
        {
            Sheets_SelectForeColor = kryptonRibbonColorButton_TextColor.SelectedColor;
            Sheets_TitleButton.ForeColor = kryptonRibbonColorButton_TextColor.SelectedColor;
            kryptonTextBox10.StateCommon.Content.Color1 = kryptonRibbonColorButton_TextColor.SelectedColor;
        }

        private void kryptonRibbonGroupComboBox_FontSize_TextUpdate(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_FontSize_TextUpdate(object sender, KeyEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void kryptonRibbonGroupComboBox_FontSize_TextUpdate(object sender, PropertyChangedEventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        public void FontReset()
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    FontStyle.Regular
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;

            }
        }

        //太字
        private void kryptonRibbonButton_Bold_Click(object sender, EventArgs e)
        {

            if (kryptonRibbonButton_Bold.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Bold
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Bold.Checked = true;


                }
            }
            else if (kryptonRibbonButton_Bold.Checked == false)
            {
                //太字ボタンをチェックをオフにする
                kryptonRibbonButton_Bold.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //斜体が有効な場合
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Italic.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem15.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem15.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem16.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem16.Checked = true;
                    }
                }
            }
        }

        //斜体
        private void kryptonRibbonButton_Italic_Click(object sender, EventArgs e)
        {

            if (kryptonRibbonButton_Italic.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Italic
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonRibbonButton_Italic.Checked = true;
                }
            }
            else if (kryptonRibbonButton_Italic.Checked == false)
            {
                //斜体ボタンをチェックをオフにする
                kryptonRibbonButton_Italic.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Bold.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem15.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem15.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem16.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem16.Checked = true;
                    }
                }
            }
        }



        private void kryptonRibbonButton_TextLine_Click(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style | FontStyle.Underline
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        //下線
        private void kryptonContextMenuItem15_CheckedChanged(object sender, EventArgs e)
        {

            if (kryptonContextMenuItem15.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Underline
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem15.Checked = true;
                }
            }
            else if (kryptonContextMenuItem15.Checked == false)
            {
                //下線メニューアイテムをチェックをオフにする
                kryptonContextMenuItem15.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Bold.Checked = true;
                    }
                }

                //斜体が有効な場合
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Italic.Checked = true;
                    }
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem16.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem16.Checked = true;
                    }
                }
            }
        }

        //打ち消し線
        private void kryptonContextMenuItem16_Click(object sender, EventArgs e)
        {

            //打ち消し線が有効な場合
            if (kryptonContextMenuItem16.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    kryptonContextMenuItem16.Checked = true;
                }
            }
            else if (kryptonContextMenuItem16.Checked == false)
            {
                //下線メニューアイテムをチェックをオフにする
                kryptonContextMenuItem16.Checked = false;
                //フォントスタイルのみ初期化
                FontReset();
                float fontSize;

                //太字が有効な場合
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Bold.Checked = true;
                    }
                }

                //斜体が有効な場合
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonRibbonButton_Italic.Checked = true;
                    }
                }

                //下線が有効な場合
                if (kryptonContextMenuItem15.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_FontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        kryptonContextMenuItem15.Checked = true;
                    }
                }
            }
        }

        private void kryptonRibbonGroupClusterButton4_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupComboBox_FontSize.Text == "8")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "9";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "9")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "9")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10.5";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10.5")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "11";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "11")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "12";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "12")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "14";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "14")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "16";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "16")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "18";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "18")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "20";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "20")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "22";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "22")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "24";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "24")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "26";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "26")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "28";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "28")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "36";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "36")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "48";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "48")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "72";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "72")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "72";
            }
        }

        private void kryptonRibbonGroupClusterButton5_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupComboBox_FontSize.Text == "8")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "8";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "9")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "8";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "9";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "10.5")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "11")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "10.5";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "12")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "11";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "14")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "12";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "16")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "14";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "18")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "16";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "20")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "18";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "22")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "20";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "24")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "22";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "26")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "24";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "28")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "26";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "36")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "28";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "48")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "36";
            }
            else if (kryptonRibbonGroupComboBox_FontSize.Text == "72")
            {
                kryptonRibbonGroupComboBox_FontSize.Text = "48";
            }
        }

        //印刷プレビュー
        private void kryptonContextMenuItem28_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Name = Sheets_TitleButton.Font.Name;
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Name = "游明朝";
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }
            GC.Collect();

            doc.PrintPreview();
        }

        private void kryptonContextMenuItem30_Click(object sender, EventArgs e)
        {
            //Wordを起動
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;

            //新しい文書を作成
            Document doc = word.Documents.Add();

            GC.Collect();
        }

        //上
        private void kryptonRibbonGroupNumericUpDown_VerticalSpace_ValueChanged(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = (int)kryptonRibbonGroupNumericUpDown_VerticalSpace.Value;
        }

        //左
        private void kryptonRibbonGroupNumericUpDown_WidthSpace_ValueChanged(object sender, EventArgs e)
        {
            Sheets_RightPanel.Height = (int)kryptonRibbonGroupNumericUpDown_WidthSpace.Value;
        }

        //下
        private void kryptonRibbonGroupNumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            Sheets_ButtomPanel.Height = (int)kryptonRibbonGroupNumericUpDown1.Value;
        }

        //右
        private void kryptonRibbonGroupNumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            Sheets_LeftPanel.Width = (int)kryptonRibbonGroupNumericUpDown2.Value;
        }

        //広い
        private void kryptonContextMenuItem21_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = 100;
            Sheets_ButtomPanel.Height = 100;
            Sheets_LeftPanel.Width = 200;
            Sheets_RightPanel.Width = 200;
            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }

        //やや狭い
        private void kryptonContextMenuItem20_Click(object sender, EventArgs e)
        {
            //上
            Sheets_TopPanel.Height = 100;
            //下
            Sheets_ButtomPanel.Height = 100;
            //右
            Sheets_LeftPanel.Width = 75;
            //左
            Sheets_RightPanel.Width = 75;

            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }

        //標準
        private void kryptonContextMenuItem18_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = 138;
            Sheets_ButtomPanel.Height = 118;
            Sheets_LeftPanel.Width = 118;
            Sheets_RightPanel.Width = 118;
            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }

        //狭い
        private void kryptonContextMenuItem19_Click(object sender, EventArgs e)
        {
            //上
            Sheets_TopPanel.Height = 50;
            //下
            Sheets_ButtomPanel.Height = 50;
            //右
            Sheets_LeftPanel.Width = 50;
            //左
            Sheets_RightPanel.Width = 50;

            //上
            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = (int)Sheets_TopPanel.Height;
            //右
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = (int)Sheets_RightPanel.Width;
            //下
            kryptonRibbonGroupNumericUpDown1.Value = (int)Sheets_ButtomPanel.Height;
            //左」
            kryptonRibbonGroupNumericUpDown2.Value = (int)Sheets_LeftPanel.Width;
        }


        private void kryptonRibbonGroupButton2_Click(object sender, EventArgs e)
        {
        }

        private void kryptonRibbonGroupButton13_Click(object sender, EventArgs e)
        {
            GC.Collect();
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //バックグラウンド上でWordを起動する
            word.Visible = false;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }

            //保存処理
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "ドキュメントファイルを保存する場所を選択 - DocuEase";
            sd.Filter = "Word 文書 (*.docx)|*.docx";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.SaveAs2(sd.FileName);
                    //Outlook連携処理
                    try
                    {
                        Microsoft.Office.Interop.Outlook.Application outlook = new Microsoft.Office.Interop.Outlook.Application();
                        MailItem mailItem = (MailItem)outlook.CreateItem(OlItemType.olMailItem);
                        mailItem.Subject = "文書送信のご案内";
                        mailItem.Attachments.Add(sd.FileName);
                        mailItem.Display(true);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Microsoft Outlook を正しく動作しませんでした。 Microsoft Outlook が正しくインストールされているか確認してください。\r\n\r\nエラー内容:\r\n" + ex.Message, "共有失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("ファイルが正しく保存されませんでした。保存するファイルの場所が適切か文書作成ソフトウェアがインストールされているか確認してください。\r\n\r\nエラー内容:\r\n" + ex.Message, "ファイル保存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            //保存を確認せず閉じる
            try
            {
                doc.Close(false);
                word.Quit();
            }
            catch { }



            GC.Collect();
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

        private void kryptonRibbonGroupButton_Tutorial_Click(object sender, EventArgs e)
        {

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                useSoftWareWindow.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            useSoftWareWindow.Show();
            kryptonRibbonGroupButton_Tutorial.Enabled = false;
            kryptonContextMenuItem12.Enabled = false;
        }

        public void kryptonContextMenuItem11_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/User233389/Document-Maker/wiki/Docuemt-Maker-%E3%83%A6%E3%83%BC%E3%82%B6%E3%83%BC%E3%82%AC%E3%82%A4%E3%83%89");
        }

        private void kryptonCommandLinkButton2_Click(object sender, EventArgs e)
        {
            ResetWarningTaskDialog resetWarningTaskDialog = new ResetWarningTaskDialog();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                resetWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            resetWarningTaskDialog.ShowDialog();

            if (resetWarningTaskDialog.DialogResult == DialogResult.Yes)
            {
                SaveFileDeleted = true;

                String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase";

                if (Directory.Exists(str))
                {
                    //ファイルを削除しトースター通知で削除したことを表示する
                    File.Delete(str + @"\SaveFile.rtf");
                    Directory.Delete(str);

                    new ToastContentBuilder()
                        .AddText("自動保存ファイルは正常に削除されました")
                        .AddText("自動保存ファイルは正しく削除されDocuEaseを終了しました。")
                        .Show();


                }
                //トースト通知のアンインストールメソッド
                ToastNotificationManagerCompat.Uninstall();
                this.Visible = false;
                IsAppRestarting = true;


                if (IsAppRestarting == true)
                {
                    //リセット前に他のウィンドウを破棄する(設定の保存はしない)
                    useSoftWareWindow.Dispose();
                    keboradShortCut.Dispose();
                    sheetsScaleDialog.Dispose();
                    //ガベージコレクション
                    GC.Collect();
                }
                //設定をリセットし保存
                Properties.Settings.Default.Reset();
                //アプリケーションを再起動
                System.Windows.Forms.Application.Restart();

            }
        }

        private void kryptonCommandLinkButton1_Click_1(object sender, EventArgs e)
        {

            Properties.Settings.Default.ShowResetDialog = true;
            Properties.Settings.Default.ShowNotepadWarningPanel = true;
            Properties.Settings.Default.Save();

            DialogResetMessagebox dialogResetMessagebox = new DialogResetMessagebox();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                dialogResetMessagebox.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            dialogResetMessagebox.ShowDialog();
        }

        private void kryptonRibbonGroupButton_Support_Click(object sender, EventArgs e)
        {
            //GitHubのWebサイトに移動
            System.Diagnostics.Process.Start("https://github.com/User233389/DocuEase");
        }

        private void kryptonTrackBar1_ValueChanged(object sender, EventArgs e)
        {
            //10の目盛りに合わせてサイズを+50上げる
            if (kryptonTrackBar1.Value == 0)
            {
                Sheets_Sheet.Size = new Size(842, 999);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                //縮小ボタンを無効化
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "10%";
                kryptonRibbonGroupComboBox1.Text = "10";
            }
            else if (kryptonTrackBar1.Value == 1)
            {
                Sheets_Sheet.Size = new Size(892, 1049);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "20%";
                kryptonRibbonGroupComboBox1.Text = "20";
            }
            else if (kryptonTrackBar1.Value == 2)
            {
                Sheets_Sheet.Size = new Size(942, 1099);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "30%";
                kryptonRibbonGroupComboBox1.Text = "30";
            }
            else if (kryptonTrackBar1.Value == 3)
            {
                Sheets_Sheet.Size = new Size(992, 1149);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "40%";
                kryptonRibbonGroupComboBox1.Text = "40";
            }
            else if (kryptonTrackBar1.Value == 4)
            {
                Sheets_Sheet.Size = new Size(1042, 1199);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "50%";
                kryptonRibbonGroupComboBox1.Text = "50";
            }
            else if (kryptonTrackBar1.Value == 5)
            {
                Sheets_Sheet.Size = new Size(1092, 1249);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "60%";
                kryptonRibbonGroupComboBox1.Text = "60";
            }
            else if (kryptonTrackBar1.Value == 6)
            {
                Sheets_Sheet.Size = new Size(1142, 1299);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "70%";
                kryptonRibbonGroupComboBox1.Text = "70";
            }
            else if (kryptonTrackBar1.Value == 7)
            {
                Sheets_Sheet.Size = new Size(1192, 1349);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "80%";
                kryptonRibbonGroupComboBox1.Text = "80";
            }
            else if (kryptonTrackBar1.Value == 8)
            {
                Sheets_Sheet.Size = new Size(1242, 1399);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "90%";
                kryptonRibbonGroupComboBox1.Text = "90";
            }
            else if (kryptonTrackBar1.Value == 9)
            {
                Sheets_Sheet.Size = new Size(1292, 1449);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "100%";
                kryptonRibbonGroupComboBox1.Text = "100";
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                Sheets_Sheet.Size = new Size(1342, 1499);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                //拡大ボタンを無効化
                kryptonButton14.Enabled = false;

                kryptonLabel42.Text = "110%";
                kryptonRibbonGroupComboBox1.Text = "110";
            }
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

        private void kryptonRibbonGroupComboBox1_TextUpdate(object sender, EventArgs e)
        {
            //10の目盛りに合わせてサイズを+50上げる
            if (kryptonRibbonGroupComboBox1.Text == "10")
            {
                Sheets_Sheet.Size = new Size(842, 999);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                //縮小ボタンを無効化
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "10%";
                kryptonTrackBar1.Value = 0;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "20")
            {
                Sheets_Sheet.Size = new Size(892, 1049);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "20%";
                kryptonTrackBar1.Value = 1;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "30")
            {
                Sheets_Sheet.Size = new Size(942, 1099);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "30%";
                kryptonTrackBar1.Value = 2;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "40")
            {
                Sheets_Sheet.Size = new Size(992, 1149);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "40%";
                kryptonTrackBar1.Value = 3;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "50")
            {
                Sheets_Sheet.Size = new Size(1042, 1199);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "50%";
                kryptonTrackBar1.Value = 4;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "60")
            {
                Sheets_Sheet.Size = new Size(1092, 1249);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "60%";
                kryptonTrackBar1.Value = 5;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "70")
            {
                Sheets_Sheet.Size = new Size(1142, 1299);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "70%";
                kryptonTrackBar1.Value = 6;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "80")
            {
                Sheets_Sheet.Size = new Size(1192, 1349);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "80%";
                kryptonTrackBar1.Value = 7;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "90")
            {
                Sheets_Sheet.Size = new Size(1242, 1399);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "90%";
                kryptonTrackBar1.Value = 8;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "100")
            {
                Sheets_Sheet.Size = new Size(1292, 1449);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;

                kryptonLabel42.Text = "100%";
                kryptonTrackBar1.Value = 9;
            }
            else if (kryptonRibbonGroupComboBox1.Text == "110")
            {
                Sheets_Sheet.Size = new Size(1342, 1499);

                kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                Sheets_Sheet.Anchor = AnchorStyles.Top;

                // 親コントロールのサイズを取得
                int parentWidth = this.ClientSize.Width;
                int parentHeight = this.ClientSize.Height;

                // パネルのサイズを取得
                int panelWidth = Sheets_Sheet.Width;
                int panelHeight = Sheets_Sheet.Height;

                // パネルの位置を中央に設定
                Sheets_Sheet.Location = new System.Drawing.Point(
                    (parentWidth - panelWidth) / 2 - 10,
                    90
                );

                Sheets_Sheet.Top = 59;

                kryptonButton15.Enabled = true;
                //拡大ボタンを無効化
                kryptonButton14.Enabled = false;

                kryptonLabel42.Text = "100%";
                kryptonTrackBar1.Value = 10;
            }

        }

        private void kryptonRibbonGroupButton2_Click_1(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton2.Checked == true)
            {
                this.WindowState = FormWindowState.Normal;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                this.AllowFormChrome = false;

                全画面モードToolStripMenuItem.Checked = true;
                kryptonRibbonGroupButton2.Checked = true;

            }
            else
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.WindowState = FormWindowState.Normal;
                this.AllowFormChrome = true;

                全画面モードToolStripMenuItem.Checked = false;
                kryptonRibbonGroupButton2.Checked = false;
            }


        }

        //戻る
        private void kryptonRibbonQATButton7_Click(object sender, EventArgs e)
        {
            if (kryptonNavigator_Workbench.SelectedPage == kryptonPage3)
            {
                kryptonNavigator_Workbench.SelectedPage = AddressTab;
                kryptonRibbonQATButton7.Enabled = true;
                前のページに移動ToolStripMenuItem.Enabled = true;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage1;
                kryptonRibbonQATButton7.Enabled = false;
                前のページに移動ToolStripMenuItem.Enabled = false;
            }
        }

        //進む
        private void kryptonRibbonQATButton8_Click(object sender, EventArgs e)
        {
            if (kryptonNavigator_Workbench.SelectedPage == kryptonPage1)
            {
                kryptonNavigator_Workbench.SelectedPage = AddressTab;
                kryptonRibbonQATButton7.Enabled = true;
                次のページに移動ToolStripMenuItem.Enabled = true;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage3;
                kryptonRibbonQATButton8.Enabled = false;
                次のページに移動ToolStripMenuItem.Enabled = false;
            }
        }

        public bool ShowReplyPanel { get; set; }
        private void kryptonNavigator_Workbench_Selected(object sender, ComponentFactory.Krypton.Navigator.KryptonPageEventArgs e)
        {
            if (menuStripPanel.Visible == true)
            {
                kryptonRibbon.MinimizedMode = false;
            }


            if (kryptonRibbon.Enabled == true)
            {
                if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
                {
                    kryptonRibbon.SelectedContext = "Address";
                    kryptonRibbon.SelectedTab = AddressTab1;

                    kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.OneNote;

                    kryptonPanel21.Height = 0;
                    kryptonPanel5.Height = 0;

                    toolStrip1.Visible = false;
                    toolStrip2.Visible = true;
                    toolStrip3.Visible = false;

                    前のページに移動ToolStripMenuItem.Enabled = true;
                    次のページに移動ToolStripMenuItem.Enabled = true;

                    コンタクトToolStripMenuItem.Visible = true;
                    メモ帳ToolStripMenuItem.Visible = false;
                }
                else if (kryptonNavigator_Workbench.SelectedPage == kryptonPage3)
                {

                    if (kryptonRibbonGroupButton_NotepadPaste.Enabled == false)
                    {
                        kryptonRibbonQATButton4.Enabled = false;
                        kryptonRibbonQATButton5.Enabled = false;
                        toolStripButton28.Enabled = false;
                        toolStripButton29.Enabled = false;
                    }
                    else if (kryptonRibbonGroupButton_NotepadPaste.Enabled == true)
                    {
                        kryptonRibbonQATButton4.Enabled = true;
                        kryptonRibbonQATButton5.Enabled = true;
                        toolStripButton28.Enabled = true;
                        toolStripButton29.Enabled = true;
                    }
                    kryptonRibbon.SelectedContext = "Notepad";
                    kryptonRibbon.SelectedTab = NotepadTab;
                    kryptonRibbonButton_Paste.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupButton_NotepadPaste.ShortcutKeys = Keys.Control | Keys.V;

                    //シート
                    kryptonRibbonButton_Bold.ShortcutKeys = Keys.None;
                    kryptonRibbonButton_Italic.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem15.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem16.ShortcutKeys = Keys.None;

                    kryptonRibbonColorButton_TextColor.ShortcutKeys = Keys.None;

                    kryptonRibbonGroupClusterButton4.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupClusterButton5.ShortcutKeys = Keys.None;

                    コピーToolStripMenuItem.ShortcutKeys = Keys.None;
                    切り取りToolStripMenuItem.ShortcutKeys = Keys.None;
                    貼り付けToolStripMenuItem.ShortcutKeys = Keys.None;
                    削除ToolStripMenuItem.ShortcutKeys = Keys.None;
                    削除ToolStripMenuItem.ShortcutKeys = Keys.None;


                    コピーToolStripMenuItem1.ShortcutKeys = Keys.Control | Keys.C;
                    切り取りToolStripMenuItem1.ShortcutKeys = Keys.Control | Keys.X;
                    貼り付けToolStripMenuItem1.ShortcutKeys = Keys.Control | Keys.V;
                    元に戻すToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.Z;
                    やり直すToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.Y;


                    //メモ
                    kryptonRibbonGroupClusterButton1.ShortcutKeys = Keys.Control | Keys.B;
                    kryptonRibbonGroupClusterButton2.ShortcutKeys = Keys.Control | Keys.I;
                    kryptonContextMenuItem35.ShortcutKeys = Keys.Control | Keys.U;
                    kryptonContextMenuItem36.ShortcutKeys = Keys.Control | Keys.T;

                    kryptonRibbonGroupColorButton2.ShortcutKeys = Keys.Control | Keys.Shift | Keys.C;
                    kryptonRibbonGroupColorButton3.ShortcutKeys = Keys.Control | Keys.Shift | Keys.M;

                    kryptonRibbonGroupClusterButton6.ShortcutKeys = Keys.Control | Keys.Shift | Keys.U;
                    kryptonRibbonGroupClusterButton7.ShortcutKeys = Keys.Control | Keys.Shift | Keys.D;

                    kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.OneNote;

                    kryptonPanel21.Height = 0;
                    kryptonPanel5.Height = 0;

                    toolStrip1.Visible = false;
                    toolStrip2.Visible = false;
                    toolStrip3.Visible = true;


                    コンタクトToolStripMenuItem.Visible = false;
                    メモ帳ToolStripMenuItem.Visible = true;

                    前のページに移動ToolStripMenuItem.Enabled = true;
                    次のページに移動ToolStripMenuItem.Enabled = false;
                }
                else
                {
                    前のページに移動ToolStripMenuItem.Enabled = false;
                    次のページに移動ToolStripMenuItem.Enabled = true;

                    toolStrip1.Visible = true;
                    toolStrip2.Visible = false;
                    toolStrip3.Visible = false;


                    コンタクトToolStripMenuItem.Visible = false;
                    メモ帳ToolStripMenuItem.Visible = false;

                    kryptonRibbon.SelectedContext = string.Empty;
                    kryptonRibbonButton_Paste.ShortcutKeys = Keys.Control | Keys.V;
                    kryptonRibbonGroupButton_NotepadPaste.ShortcutKeys = Keys.None;

                    kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

                    Sheets_Sheet.Anchor = AnchorStyles.Top;

                    // 親コントロールのサイズを取得
                    int parentWidth = this.ClientSize.Width;
                    int parentHeight = this.ClientSize.Height;

                    // パネルのサイズを取得
                    int panelWidth = Sheets_Sheet.Width;
                    int panelHeight = Sheets_Sheet.Height;

                    // パネルの位置を中央に設定
                    Sheets_Sheet.Location = new System.Drawing.Point(
                        (parentWidth - panelWidth) / 2 - 10,
                        90
                    );


                    コピーToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.C;
                    切り取りToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.X;
                    貼り付けToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.V;
                    削除ToolStripMenuItem.ShortcutKeys = Keys.Delete;


                    コピーToolStripMenuItem1.ShortcutKeys = Keys.None;
                    切り取りToolStripMenuItem1.ShortcutKeys = Keys.None;
                    貼り付けToolStripMenuItem1.ShortcutKeys = Keys.None;
                    元に戻すToolStripMenuItem.ShortcutKeys = Keys.None;
                    やり直すToolStripMenuItem.ShortcutKeys = Keys.None;

                    //シート
                    kryptonRibbonButton_Bold.ShortcutKeys = Keys.Control | Keys.B;
                    kryptonRibbonButton_Italic.ShortcutKeys = Keys.Control | Keys.I;
                    kryptonContextMenuItem15.ShortcutKeys = Keys.Control | Keys.U;
                    kryptonContextMenuItem16.ShortcutKeys = Keys.Control | Keys.T;

                    kryptonRibbonColorButton_TextColor.ShortcutKeys = Keys.Control | Keys.Shift | Keys.C;

                    kryptonRibbonGroupClusterButton4.ShortcutKeys = Keys.Control | Keys.Shift | Keys.U;
                    kryptonRibbonGroupClusterButton5.ShortcutKeys = Keys.Control | Keys.Shift | Keys.D;

                    //メモ
                    kryptonRibbonGroupClusterButton1.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupClusterButton2.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem35.ShortcutKeys = Keys.None;
                    kryptonContextMenuItem36.ShortcutKeys = Keys.None;

                    kryptonRibbonGroupColorButton2.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupColorButton3.ShortcutKeys = Keys.None;

                    kryptonRibbonGroupClusterButton6.ShortcutKeys = Keys.None;
                    kryptonRibbonGroupClusterButton7.ShortcutKeys = Keys.None;

                    if (kryptonRibbonGroupButton16.Checked == true | 置換ToolStripMenuItem.Checked == true)
                    {
                        kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.SlantOutsizeFar;
                        kryptonPanel21.Height = 36;
                        kryptonRibbonGroupButton16.Checked = true;
                        置換ToolStripMenuItem.Checked = true;
                    }
                    else if (kryptonRibbonGroupButton16.Checked == false | 置換ToolStripMenuItem.Checked == false)
                    {
                        kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.OneNote;
                        kryptonPanel21.Height = 0;
                        kryptonRibbonGroupButton16.Checked = false;
                        置換ToolStripMenuItem.Checked = false;
                    }


                    kryptonPanel5.Height = 26;
                }


            }

            Sheets_Sheet.Top = 59;

            if (kryptonNavigator_Workbench.SelectedPage == kryptonPage1)
            {
                kryptonRibbonQATButton7.Enabled = false;
                kryptonRibbonQATButton8.Enabled = true;

                kryptonRibbonQATButton4.Enabled = false;
                kryptonRibbonQATButton5.Enabled = false;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == AddressTab)
            {
                kryptonRibbonQATButton7.Enabled = true;
                kryptonRibbonQATButton8.Enabled = true;

                kryptonRibbonQATButton4.Enabled = false;
                kryptonRibbonQATButton5.Enabled = false;
            }
            else if (kryptonNavigator_Workbench.SelectedPage == kryptonPage3)
            {
                kryptonRibbonQATButton7.Enabled = true;
                kryptonRibbonQATButton8.Enabled = false;

                kryptonRibbonQATButton4.Enabled = true;
                kryptonRibbonQATButton5.Enabled = true;
            }
        }


        //メモ帳
        //元に戻す
        private void kryptonRibbonQATButton4_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Undo();
        }

        //やり直す
        private void kryptonRibbonQATButton5_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Redo();
        }

        //貼り付け
        private void kryptonRibbonGroupButton_NotepadPaste_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Paste();
        }

        private void kryptonRibbonGroupButton_NotepadCopy_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Copy();
        }

        private void kryptonRibbonGroupButton_NotepadCut_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Cut();
        }


        public int rtbLangth { get; set; }
        public int rtbStart { get; set; }

        //フォント変更
        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, EventArgs e)
        {

            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
            // 現在のフォント名を変更する
            Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                kryptonRibbonGroupComboBox_NotepadFont.Text,
                Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style
            );
            toolStripComboBox3.Text = kryptonRibbonGroupComboBox_NotepadFont.Text;

        }


        //フォントサイズを変更
        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, EventArgs e)
        {

            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                    fontSize,
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style
                );
            }
            toolStripComboBox4.Text = kryptonRibbonGroupComboBox_NotepadFontSize.Text;

        }

        private void Notepads_kryptonRichTextBox_Notepad_SelectionChanged(object sender, EventArgs e)
        {
            //文字選択数取得
            rtbLangth = Notepads_kryptonRichTextBox_Notepad.SelectionLength;
            rtbStart = Notepads_kryptonRichTextBox_Notepad.SelectionStart;

            //箇条書き確認
            if (Notepads_kryptonRichTextBox_Notepad.SelectionBullet == true)
            {
                kryptonRibbonGroupClusterButton13.Checked = true;
                toolStripButton39.Checked = true;
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionBullet == false)
            {
                kryptonRibbonGroupClusterButton13.Checked = false;
                toolStripButton39.Checked = false;
            }

            try
            {
                //文字色の確認
                kryptonRibbonGroupColorButton2.SelectedColor = Notepads_kryptonRichTextBox_Notepad.SelectionColor;
                //マーカー色の確認
                kryptonRibbonGroupColorButton3.SelectedColor = Notepads_kryptonRichTextBox_Notepad.SelectionBackColor;
                //フォントスタイルの確認
                //太字
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                    toolStripButton18.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton1.Checked = false;
                    toolStripButton18.Checked = false;
                }

                //斜体
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                    toolStripButton19.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton2.Checked = false;
                    toolStripButton19.Checked = false;
                }

                //下線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                    toolStripMenuItem2.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem35.Checked = false;
                    toolStripMenuItem2.Checked = false;
                }

                //打ち消し線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                    toolStripMenuItem3.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem35.Checked = false;
                    toolStripMenuItem3.Checked = false;
                }

                // 選択範囲のフォントを取得（同一フォントでない場合は null を返す）
                System.Drawing.Font selFont = Notepads_kryptonRichTextBox_Notepad.SelectionFont;

                if (selFont != null)
                {
                    // 選択範囲が単一フォントの場合はそのまま反映
                    kryptonRibbonGroupComboBox_NotepadFont.Text = selFont.Name;
                    kryptonRibbonGroupComboBox_NotepadFontSize.Text = selFont.Size.ToString();
                }
                else
                {
                    // フォントが混在している場合は、選択範囲の先頭文字のフォントを取得して表示（元の選択は復元）
                    int selStart = Notepads_kryptonRichTextBox_Notepad.SelectionStart;
                    int selLen = Notepads_kryptonRichTextBox_Notepad.SelectionLength;

                    if (selStart < Notepads_kryptonRichTextBox_Notepad.TextLength && selLen > 0)
                    {
                        // 一時的に先頭1文字を選択してフォントを調べる
                        Notepads_kryptonRichTextBox_Notepad.Select(selStart, 1);
                        System.Drawing.Font firstCharFont = Notepads_kryptonRichTextBox_Notepad.SelectionFont;

                        // 元の選択範囲を復元
                        Notepads_kryptonRichTextBox_Notepad.Select(selStart, selLen);

                        if (firstCharFont != null)
                        {
                            // 「混在」表記を付けて先頭のフォント情報を表示
                            kryptonRibbonGroupComboBox_NotepadFont.Text = firstCharFont.Name + " (混在)";
                            kryptonRibbonGroupComboBox_NotepadFontSize.Text = firstCharFont.Size.ToString();
                            return;
                        }
                    }

                    // 選択なしやフォント情報が取れない場合は空にする
                    kryptonRibbonGroupComboBox_NotepadFont.Text = string.Empty;
                    kryptonRibbonGroupComboBox_NotepadFontSize.Text = string.Empty;
                }
            }
            catch (System.Exception)
            {
                kryptonRibbonGroupComboBox_NotepadFont.Text = "(混在)";
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "(混在)";
            }


        }

        private void Notepads_kryptonRichTextBox_Notepad_MouseUp(object sender, MouseEventArgs e)
        {



        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDown(object sender, EventArgs e)
        {

        }

        private void Notepads_kryptonRichTextBox_Notepad_FontChanged(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, KeyEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, KeyPressEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFont_DropDownClosed(object sender, PropertyChangedEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, KeyEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, KeyPressEventArgs e)
        {

        }

        private void kryptonRibbonGroupComboBox_NotepadFontSize_DropDownClosed(object sender, PropertyChangedEventArgs e)
        {

        }
        public void FontReset2()
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                    fontSize,
                    FontStyle.Regular
                );
            }
        }

        //太字
        private void kryptonRibbonGroupClusterButton1_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupClusterButton1.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonRibbonGroupClusterButton1.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }

            }
            else if (kryptonRibbonGroupClusterButton1.Checked == false)
            {
                kryptonRibbonGroupClusterButton1.Checked = false;
                FontReset2();

                //斜体が有効な場合
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    kryptonRibbonGroupClusterButton2.Checked = true;
                }

                //下線が有効な場合
                if (kryptonContextMenuItem35.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    kryptonContextMenuItem35.Checked = true;
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem36.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    kryptonContextMenuItem36.Checked = true;
                }
            }

        }

        //斜体
        private void kryptonRibbonGroupClusterButton2_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupClusterButton2.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonRibbonGroupClusterButton2.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }
            }
            else if (kryptonRibbonGroupClusterButton2.Checked == false)
            {
                FontReset2();

                //太字が有効な場合
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    kryptonRibbonGroupClusterButton1.Checked = true;
                }

                //下線が有効な場合
                if (kryptonContextMenuItem35.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    kryptonContextMenuItem35.Checked = true;
                }

                //打ち消し線が有効な場合
                if (kryptonContextMenuItem36.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    kryptonContextMenuItem36.Checked = true;
                }

            }
        }

        //下線
        private void kryptonContextMenuItem35_Click(object sender, EventArgs e)
        {
            if (kryptonContextMenuItem35.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonContextMenuItem35.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }


            }
            else if (kryptonContextMenuItem35.Checked == false)
            {
                FontReset2();
                //太字
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    kryptonRibbonGroupClusterButton1.Checked = true;
                }

                //斜体
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    kryptonRibbonGroupClusterButton2.Checked = true;
                }



                //打ち消し線
                if (kryptonContextMenuItem36.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    kryptonContextMenuItem36.Checked = true;
                }
            }
        }

        //打ち消し線
        private void kryptonContextMenuItem36_Click(object sender, EventArgs e)
        {
            if (kryptonContextMenuItem36.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                    );
                }

                //完了後他のフォントスタイルを確認
                kryptonContextMenuItem36.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
            }
            else if (kryptonContextMenuItem36.Checked == false)
            {
                FontReset2();
                //太字
                if (kryptonRibbonGroupClusterButton1.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    kryptonRibbonGroupClusterButton1.Checked = true;
                }

                //斜体
                if (kryptonRibbonGroupClusterButton2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    kryptonRibbonGroupClusterButton2.Checked = true;
                }

                //下線
                if (kryptonContextMenuItem35.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    kryptonContextMenuItem35.Checked = true;
                }

            }
        }

        //文字色
        private void kryptonRibbonGroupColorButton2_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionColor = kryptonRibbonGroupColorButton2.SelectedColor;
        }

        //マーカー色
        private void kryptonRibbonGroupColorButton3_SelectedColorChanged(object sender, ComponentFactory.Krypton.Toolkit.ColorEventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionBackColor = kryptonRibbonGroupColorButton3.SelectedColor;
        }

        private void kryptonRibbonGroupButton_NotepadSaveAs_Click(object sender, EventArgs e)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "メモを保存する場所を選択...";
            sd.Filter = "リッチテキストファイル (*.rtf)|*.rtf|書式なしテキストファイル(*.txt)|*.txt";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                //rtfファイルだった場合
                if (sd.FilterIndex == 1)
                {
                    Notepads_kryptonRichTextBox_Notepad.SaveFile(sd.FileName);
                }
                //txtファイルだった場合
                else if (sd.FilterIndex == 2)
                {
                    StreamWriter writer = new StreamWriter(sd.FileName);
                    string str = Notepads_kryptonRichTextBox_Notepad.Text;
                    writer.WriteLine(str);
                    writer.Close();
                    writer.Dispose();
                }
            }
        }



        private void kryptonRibbonGroup2_DialogBoxLauncherClick(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            KryptonFontDialog fd = new KryptonFontDialog();
            fd.DisplayExtendedColorsButton = true;
            fd.Font = Notepads_kryptonRichTextBox_Notepad.SelectionFont;
            fd.ShowColor = true;
            fd.Color = Notepads_kryptonRichTextBox_Notepad.SelectionColor;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    fd.Font.Name,
                    fd.Font.Size,
                    fd.Font.Style
                );

                //フォント名を確認する
                kryptonRibbonGroupComboBox_NotepadFont.Text = fd.Font.Name;
                //フォントサイズを確認する
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = fd.Font.Size.ToString();
                //フォントスタイルを確認する
                //太字
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    kryptonRibbonGroupClusterButton1.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton1.Checked = false;
                }
                //斜体
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    kryptonRibbonGroupClusterButton2.Checked = true;
                }
                else
                {
                    kryptonRibbonGroupClusterButton2.Checked = false;
                }
                //下線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    kryptonContextMenuItem35.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem35.Checked = false;
                }
                //打ち消し線
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    kryptonContextMenuItem36.Checked = true;
                }
                else
                {
                    kryptonContextMenuItem36.Checked = false;
                }
                //文字色を確認する
                kryptonRibbonGroupColorButton2.SelectedColor = fd.Color;

            }
        }

        //フォントサイズを上げる
        private void kryptonRibbonGroupClusterButton6_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 8)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "9";
                toolStripComboBox4.Text = "9";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 9)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "10";
                toolStripComboBox4.Text = "10";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 10)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "11";
                toolStripComboBox4.Text = "11";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 11)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "12";
                toolStripComboBox4.Text = "12";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 12)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "14";
                toolStripComboBox4.Text = "14";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 14)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "16";
                toolStripComboBox4.Text = "16";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 16)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "18";
                toolStripComboBox4.Text = "18";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 18)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "20";
                toolStripComboBox4.Text = "20";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 20)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "22";
                toolStripComboBox4.Text = "22";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 22)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "24";
                toolStripComboBox4.Text = "24";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 24)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "26";
                toolStripComboBox4.Text = "26";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 26)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "28";
                toolStripComboBox4.Text = "28";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 28)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "36";
                toolStripComboBox4.Text = "36";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 36)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "48";
                toolStripComboBox4.Text = "48";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 48)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "72";
                toolStripComboBox4.Text = "72";
            }
        }

        //フォントサイズを下げる
        private void kryptonRibbonGroupClusterButton7_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 9)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "8";
                toolStripComboBox4.Text = "8";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 10)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "9";
                toolStripComboBox4.Text = "9";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 11)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "10";
                toolStripComboBox4.Text = "10";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 12)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "11";
                toolStripComboBox4.Text = "11";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 14)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "12";
                toolStripComboBox4.Text = "12";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 16)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "14";
                toolStripComboBox4.Text = "14";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 18)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "16";
                toolStripComboBox4.Text = "16";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 20)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "18";
                toolStripComboBox4.Text = "18";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 22)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "20";
                toolStripComboBox4.Text = "20";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 24)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "22";
                toolStripComboBox4.Text = "22";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 26)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "24";
                toolStripComboBox4.Text = "24";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 28)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "26";
                toolStripComboBox4.Text = "26";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 36)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "28";
                toolStripComboBox4.Text = "28";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 48)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "36";
                toolStripComboBox4.Text = "36";
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size == 72)
            {
                kryptonRibbonGroupComboBox_NotepadFontSize.Text = "48";
                toolStripComboBox4.Text = "48";
            }
        }

        private void kryptonRibbonGroupButton20_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Title = "画像ファイルを選択...";
            od.Filter = "PNGファイル(*.png)|*.png|JPEGファイル(*.jpeg)|*.jpeg|JPGファイル(*.jpg)|*.jpg";
            if (od.ShowDialog() == DialogResult.OK)
            {
                Clipboard.SetImage(Image.FromFile(od.FileName));
                Notepads_kryptonRichTextBox_Notepad.Paste();
            }
        }


        private void kryptonRibbonGroupButton15_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton15.Checked == true)
            {
                kryptonRibbonGroup13.Visible = true;
            }
            else
            {
                kryptonRibbonGroup13.Visible = false;
            }
        }

        private void Notepads_kryptonRichTextBox_Notepad_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void kryptonCheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (kryptonCheckBox1.Checked == true)
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                //下記のように西暦ではなく和暦として表示するように設定する
                culturejp.DateTimeFormat.Calendar = new JapaneseCalendar();
                Sheets_DateLabel.Text = date.ToString("ggy年M月d日", culturejp);
            }
            else
            {
                DateTime date = kryptonDateTimePicker1.Value.Date;
                CultureInfo culturejp = new CultureInfo("ja-Jp");
                Sheets_DateLabel.Text = date.ToString("yyyy年M月d日");
            }
        }

        private void kryptonContextMenuItem53_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Focused == true)
            {
                kryptonTextBox11.Focus();
                kryptonTextBox11.SelectAll();
            }
            else if (kryptonTextBox1.Focused == true)
            {
                kryptonTextBox1.Focus();
                kryptonTextBox1.SelectAll();
            }
            else if (kryptonComboBox10.Focused == true)
            {
                kryptonComboBox10.Focus();
                kryptonComboBox10.SelectAll();
            }
            else if (kryptonTextBox2.Focused == true)
            {
                kryptonTextBox2.Focus();
                kryptonTextBox2.SelectAll();
            }
            else if (kryptonTextBox3.Focused == true)
            {
                kryptonTextBox3.Focus();
                kryptonTextBox3.SelectAll();
            }
            else if (kryptonTextBox4.Focused == true)
            {
                kryptonTextBox4.Focus();
                kryptonTextBox4.SelectAll();
            }
            else if (kryptonTextBox5.Focused == true)
            {
                kryptonTextBox5.Focus();
                kryptonTextBox5.SelectAll();
            }
            else if (kryptonComboBox9.Focused == true)
            {
                kryptonComboBox9.Focus();
                kryptonComboBox9.SelectAll();
            }
            else if (kryptonTextBox6.Focused == true)
            {
                kryptonTextBox6.Focus();
                kryptonTextBox6.SelectAll();
            }
            else if (kryptonTextBox7.Focused == true)
            {
                kryptonTextBox7.Focus();
                kryptonTextBox7.SelectAll();
            }
            else if (kryptonComboBox8.Focused == true)
            {
                kryptonComboBox8.Focus();
                kryptonComboBox8.SelectAll();
            }
            else if (kryptonComboBox6.Focused == true)
            {
                kryptonTextBox6.Focus();
                kryptonTextBox6.SelectAll();
            }
            else if (kryptonTextBox14.Focused == true)
            {
                kryptonTextBox14.Focus();
                kryptonTextBox14.SelectAll();
            }
            else if (kryptonTextBox8.Focused == true)
            {
                kryptonTextBox8.Focus();
                kryptonTextBox8.SelectAll();
            }
            else if (kryptonComboBox7.Focused == true)
            {
                kryptonComboBox7.Focus();
                kryptonComboBox7.SelectAll();
            }
            else if (kryptonTextBox9.Focused == true)
            {
                kryptonTextBox9.Focus();
                kryptonTextBox9.SelectAll();
            }
            else if (kryptonTextBox15.Focused == true)
            {
                kryptonTextBox15.Focus();
                kryptonTextBox15.SelectAll();
            }
            else if (kryptonTextBox10.Focused == true)
            {
                kryptonTextBox10.Focus();
                kryptonTextBox10.SelectAll();
            }
            else if (kryptonComboBox2.Focused == true)
            {
                kryptonComboBox2.Focus();
                kryptonComboBox2.SelectAll();
            }
            else if (kryptonComboBox11.Focused == true)
            {
                kryptonComboBox11.Focus();
                kryptonComboBox11.SelectAll();
            }
            else if (kryptonComboBox3.Focused == true)
            {
                kryptonComboBox3.Focus();
                kryptonComboBox3.SelectAll();
            }
            else if (kryptonComboBox4.Focused == true)
            {
                kryptonComboBox4.Focus();
                kryptonComboBox4.SelectAll();
            }
            else if (kryptonTextBox12.Focused == true)
            {
                kryptonTextBox12.Focus();
                kryptonTextBox12.SelectAll();
            }
            else if (kryptonComboBox5.Focused == true)
            {
                kryptonComboBox5.Focus();
                kryptonComboBox5.SelectAll();
            }
            else if (kryptonTextBox13.Focused == true)
            {
                kryptonTextBox13.Focus();
                kryptonTextBox13.SelectAll();
            }
        }

        //発行元部署名
        private void kryptonContextMenuItem57_Click(object sender, EventArgs e)
        {
            kryptonTextBox11.Focus();
        }

        //発行番号
        private void kryptonContextMenuItem58_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown1.Focus();
        }

        //日付
        private void kryptonContextMenuItem79_Click(object sender, EventArgs e)
        {
            kryptonDateTimePicker1.Focus();
        }

        //組織および会社名
        private void kryptonContextMenuItem59_Click(object sender, EventArgs e)
        {
            kryptonTextBox1.Focus();
        }

        //肩書きと氏名
        private void kryptonContextMenuItem61_Click(object sender, EventArgs e)
        {
            kryptonComboBox10.Focus();
        }

        //組織および会社名
        private void kryptonContextMenuItem62_Click(object sender, EventArgs e)
        {
            kryptonTextBox3.Focus();
        }

        //所在地
        private void kryptonContextMenuItem63_Click(object sender, EventArgs e)
        {
            kryptonTextBox4.Focus();
        }

        //建物名と階数
        private void kryptonContextMenuItem64_Click(object sender, EventArgs e)
        {
            kryptonTextBox5.Focus();
        }

        //肩書きと氏名
        private void kryptonContextMenuItem65_Click(object sender, EventArgs e)
        {
            kryptonComboBox9.Focus();
        }

        //メールアドレス
        private void kryptonContextMenuItem66_Click(object sender, EventArgs e)
        {
            kryptonTextBox7.Focus();
        }

        //電話番号
        private void kryptonContextMenuItem67_Click(object sender, EventArgs e)
        {
            kryptonComboBox6.Focus();
        }

        //Fax番号
        private void kryptonContextMenuItem68_Click(object sender, EventArgs e)
        {
            kryptonComboBox7.Focus();
        }

        //表題名
        private void kryptonContextMenuItem69_Click(object sender, EventArgs e)
        {
            kryptonTextBox10.Focus();
        }

        //月
        private void kryptonContextMenuItem70_Click(object sender, EventArgs e)
        {
            kryptonComboBox1.Focus();
        }

        //頭語
        private void kryptonContextMenuItem71_Click(object sender, EventArgs e)
        {
            kryptonComboBox2.Focus();
        }

        //候文
        private void kryptonContextMenuItem72_Click(object sender, EventArgs e)
        {
            kryptonComboBox11.Focus();
        }

        //感謝のあいさつ1
        private void kryptonContextMenuItem73_Click(object sender, EventArgs e)
        {
            kryptonComboBox3.Focus();
        }

        //感謝のあいさつ2
        private void kryptonContextMenuItem74_Click(object sender, EventArgs e)
        {
            kryptonComboBox4.Focus();
        }

        //結語
        private void kryptonContextMenuItem75_Click(object sender, EventArgs e)
        {
            kryptonComboBox5.Focus();
        }


        //内容文
        private void kryptonContextMenuItem76_Click(object sender, EventArgs e)
        {
            kryptonTextBox12.Focus();
        }

        //記し書き文
        private void kryptonContextMenuItem77_Click(object sender, EventArgs e)
        {
            kryptonTextBox13.Focus();
        }

        //ToDo
        private void kryptonContextMenuItem37_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・ToDo\r\n1.\r\n2.\r\n3.";
            kryptonRibbonGroupButton5.TextLine1 = "・ToDo";
        }

        //やることリスト
        private void kryptonContextMenuItem38_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・やることリスト\r\n1.\r\n2.\r\n3.";
            kryptonRibbonGroupButton5.TextLine1 = "・やることリスト";
        }

        //宛先
        private void kryptonContextMenuItem39_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・宛先";
            kryptonRibbonGroupButton5.TextLine1 = "・宛先";
        }

        //発信者
        private void kryptonContextMenuItem40_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・発信者";
            kryptonRibbonGroupButton5.TextLine1 = "・発信者";
        }

        //表題
        private void kryptonContextMenuItem41_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・表題";
            kryptonRibbonGroupButton5.TextLine1 = "・表題";
        }

        //内容と記し書き
        private void kryptonContextMenuItem42_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・内容\r\n\r\n・記し書き";
            kryptonRibbonGroupButton5.TextLine1 = "・内容と記し書き";
        }

        //概要
        private void kryptonContextMenuItem43_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・概要";
            kryptonRibbonGroupButton5.TextLine1 = "・概要";
        }

        //要点
        private void kryptonContextMenuItem44_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・要点";
            kryptonRibbonGroupButton5.TextLine1 = "・要点";
        }

        //注意
        private void kryptonContextMenuItem45_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・注意";
            kryptonRibbonGroupButton5.TextLine1 = "・注意";
        }

        private void kryptonRibbonGroupButton5_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton5.TextLine1 == "・ToDo")
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・ToDo\r\n1.\r\n2.\r\n3.";
            }
            else if (kryptonRibbonGroupButton5.TextLine1 == "・やることリスト")
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・やることリスト\r\n1.\r\n2.\r\n3.";
            }
            else if (kryptonRibbonGroupButton5.TextLine1 == "・内容と記し書き")
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・内容\r\n\r\n・記し書き";
            }
            else
            {
                Notepads_kryptonRichTextBox_Notepad.Text += "\r\n" + kryptonRibbonGroupButton5.TextLine1;
            }

        }

        //最高
        private void kryptonContextMenuItem46_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・最高";
            kryptonRibbonGroupButton8.TextLine1 = "・最高";
        }

        //高
        private void kryptonContextMenuItem47_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・高";
            kryptonRibbonGroupButton8.TextLine1 = "・高";
        }

        //中
        private void kryptonContextMenuItem48_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・中";
            kryptonRibbonGroupButton8.TextLine1 = "・中";
        }

        //小
        private void kryptonContextMenuItem49_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・小";
            kryptonRibbonGroupButton8.TextLine1 = "・小";
        }

        //緊急
        private void kryptonContextMenuItem50_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・緊急";
            kryptonRibbonGroupButton8.TextLine1 = "・緊急";
        }

        //要確認
        private void kryptonContextMenuItem51_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・要確認";
            kryptonRibbonGroupButton8.TextLine1 = "・要確認";
        }

        //状態
        private void kryptonContextMenuItem52_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n・状態";
            kryptonRibbonGroupButton8.TextLine1 = "・状態";
        }

        private void kryptonRibbonGroupButton8_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Text += "\r\n" + kryptonRibbonGroupButton5.TextLine1;
        }

        private void kryptonContextMenuItem92_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectAll();
        }

        //ショートカット
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.F9)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage1;
            }
            else if (e.Control && e.KeyCode == Keys.F11)
            {
                kryptonNavigator_Workbench.SelectedPage = AddressTab;
            }
            else if (e.Control && e.KeyCode == Keys.F12)
            {
                kryptonNavigator_Workbench.SelectedPage = kryptonPage3;
            }

            if (e.Control && e.Shift && e.KeyCode == Keys.W)
            {
                kryptonLabel1.Text = "出力中...";
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                word.Visible = true;
                Document doc = word.Documents.Add();


                //外枠の余白を設定
                doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
                doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
                doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
                doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

                foreach (Range range in doc.StoryRanges)
                {
                    range.Font.Size = 10; // フォントサイズを10に設定
                }

                //発行番号
                if (Sheets_NumberLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                    paragraph1.Range.Text = Sheets_NumberLabel.Text;
                    paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph1.Range.InsertParagraphAfter();
                }
                //日付
                if (Sheets_DateLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                    paragraph2.Range.Text = Sheets_DateLabel.Text;
                    paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph2.Range.InsertParagraphAfter();
                }
                //相手先会社名
                if (Sheets_AddressCompanyLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                    paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                    paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph3.Range.InsertParagraphAfter();
                }
                //相手先氏名
                if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                    paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                    paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph4.Range.InsertParagraphAfter();
                }
                //発信者会社名
                if (Sheets_CallerCompanyLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                    paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                    paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph5.Range.InsertParagraphAfter();
                }
                //発信者所在地
                if (Sheets_CallerLocationLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                    paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                    paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph6.Range.InsertParagraphAfter();
                }
                //発信者建物名と階数
                if (Sheets_BuildingNameLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                    paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                    paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph7.Range.InsertParagraphAfter();
                }
                //発信者氏名
                if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                    paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                    paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph8.Range.InsertParagraphAfter();
                }
                //メールアドレス
                if (Sheets_CallerMallAddressLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                    paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                    paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph9.Range.InsertParagraphAfter();
                }
                //電話番号
                if (Sheets_CallerTelLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                    paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                    paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph10.Range.InsertParagraphAfter();
                }
                //Fax番号
                if (Sheets_CallerFaxTelLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                    paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                    paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph11.Range.InsertParagraphAfter();
                }
                //表題
                if (Sheets_TitleButton.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                    if (kryptonRibbonButton_Bold.Checked == true)
                    {
                        paragraph12.Range.Bold = 1;
                    }
                    if (kryptonRibbonButton_Italic.Checked == true)
                    {
                        paragraph12.Range.Italic = 1;
                    }
                    if (kryptonContextMenuItem15.Checked == true)
                    {
                        paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                    }
                    if (kryptonContextMenuItem16.Checked == true)
                    {
                        paragraph12.Range.Font.StrikeThrough = 1;
                    }
                    paragraph12.Range.Font.Name = Sheets_TitleButton.Font.Name;
                    paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                    paragraph12.Range.Text = Sheets_TitleButton.Text;
                    paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                    paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph12.Range.InsertParagraphAfter();

                }
                //あいさつ文
                if (Sheets_ContentLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                    paragraph13.Range.Bold = 0;
                    paragraph13.Range.Italic = 0;
                    paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                    paragraph13.Range.Font.StrikeThrough = 0;
                    paragraph13.Range.Font.Name = "游明朝";
                    paragraph13.Range.Font.Size = 10;
                    paragraph13.Range.Text = Sheets_ContentLabel.Text;
                    paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                    paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph13.Range.InsertParagraphAfter();
                }
                //内容
                try
                {
                    int LinesCount = 0;
                    while (true)
                    {
                        Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                        W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                        W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                        W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                        W_Contents.Range.InsertParagraphAfter();
                        LinesCount = LinesCount + 1;
                        if (LinesCount == kryptonTextBox12.Lines.Length)
                        {
                            break;
                        }
                    }
                }
                catch { }
                //結語
                //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
                if (kryptonRibbonGroupCheckBox1.Checked != true)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                    paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                    paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph14.Range.InsertParagraphAfter();
                }
                kryptonLabel1.Text = "出力完了";
                stausUpdate();
                //記
                if (Sheet_NoteLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                    paragraph15.Range.Text = Sheet_NoteLabel.Text;
                    paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph15.Range.InsertParagraphAfter();
                }
                //記し書き
                try
                {
                    int LinesCount2 = 0;
                    while (true)
                    {
                        Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                        W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                        W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                        W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                        W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                        W_Contents.Range.InsertParagraphAfter();
                        LinesCount2 = LinesCount2 + 1;
                        if (LinesCount2 == kryptonTextBox13.Lines.Length)
                        {
                            break;
                        }
                    }
                }
                catch { }
                //以上
                if (Sheets_EndLabel.Text != string.Empty)
                {
                    Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                    paragraph16.Range.Text = Sheets_EndLabel.Text;
                    paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    paragraph16.Range.InsertParagraphAfter();
                }
                GC.Collect();

                doc.PrintPreview();
            }
        }

        private void kryptonNavigator_Workbench_KeyDown(object sender, KeyEventArgs e)
        {

        }


        private void buttonSpecAppMenu2_Click(object sender, EventArgs e)
        {

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                keboradShortCut.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            keboradShortCut.Show();
        }
        public string flName { get; set; }
        public string loaction { get; set; }

        public string ContactEmailAddress { get; set; }
        public string MailAddress_User { get; set; }
        public string MailAddress_Domain { get; set; }

        public string PhoneNumber1 { get; set; }
        public string PhoneNumber2 { get; set; }
        public string PhoneNumber3 { get; set; }

        public string FaxNumber1 { get; set; }
        public string FaxNumber2 { get; set; }
        public string FacNumber3 { get; set; }

        private void kryptonListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            kryptonButton6.Show();
            // Outlookアプリケーションのインスタンスを取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            // 検索したい名前
            string targetName = kryptonListBox3.SelectedItem.ToString();

            // 連絡先を検索
            Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
            Microsoft.Office.Interop.Outlook.ContactItem contact = contactItems.Find($"[FullName] = '{targetName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

            if (contact != null)
            {

                loaction = contact.BusinessAddress;

                Address_NameLabel.Text = kryptonListBox3.SelectedItem.ToString();
                kryptonCheckBox10.Text = "所在地:〒" + contact.BusinessAddress;
                kryptonCheckBox11.Text = "メールアドレス:" + contact.Email1Address;
                ContactEmailAddress = contact.Email1Address;
                kryptonCheckBox12.Text = "会社電話番号:" + contact.BusinessTelephoneNumber;
                kryptonCheckBox13.Text = "会社Fax番号:" + contact.BusinessFaxNumber;

                //メールアドレス
                // 変更箇所: kryptonListBox3_SelectedIndexChanged 内の contact.Email1Address を分割してユーザー名とドメイン名を文字列に格納する処理
                // 以下を該当メソッド内の該当行（kryptonRadioButton4.Text = ... の代わり）に置き換えてください。

                // フルアドレスをプロパティに保管（既存プロパティを活用）
                ContactEmailAddress = contact.Email1Address ?? string.Empty;

                // ユーザー名とドメイン名を分割して格納
                MailAddress_User = string.Empty;
                MailAddress_Domain = string.Empty;
                string email = ContactEmailAddress.Trim();

                if (!string.IsNullOrEmpty(email))
                {
                    int at = email.IndexOf('@');
                    if (at > 0 && at < email.Length - 1)
                    {
                        MailAddress_User = email.Substring(0, at);
                        MailAddress_Domain = email.Substring(at + 1);
                    }
                    else
                    {
                        // @ が無い・不正な形式の場合は全体をユーザー部として扱う
                        MailAddress_User = email;
                        MailAddress_Domain = string.Empty;
                    }
                }

                //電話番号
                // 会社電話番号を最大3つのパートに分割して格納
                PhoneNumber1 = PhoneNumber2 = PhoneNumber3 = string.Empty;
                string tel = (contact.BusinessTelephoneNumber ?? string.Empty).Trim();

                if (!string.IsNullOrEmpty(tel))
                {
                    // 数字以外で分割（ハイフンやスペース、括弧などを区切りにする）
                    string[] rawParts = System.Text.RegularExpressions.Regex.Split(tel, @"\D+");
                    System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
                    foreach (var p in rawParts)
                    {
                        if (!string.IsNullOrEmpty(p)) parts.Add(p);
                    }

                    if (parts.Count >= 3)
                    {
                        PhoneNumber1 = parts[0];
                        PhoneNumber2 = parts[1];
                        PhoneNumber3 = parts[2];
                    }
                    else if (parts.Count == 2)
                    {
                        PhoneNumber1 = parts[0];
                        PhoneNumber2 = parts[1];
                        PhoneNumber3 = string.Empty;
                    }
                    else if (parts.Count == 1)
                    {
                        // 桁数に応じて分割（簡易フォールバック）
                        string digits = parts[0];
                        int len = digits.Length;
                        if (len >= 7)
                        {
                            // 末尾4桁を最後のパートに確保し、先頭を残りで分割
                            int last = 4;
                            int first = Math.Max(2, len - 7); // 最低2桁を先頭に
                            int middle = len - first - last;
                            if (first > 0 && middle >= 0)
                            {
                                PhoneNumber1 = digits.Substring(0, first);
                                if (middle > 0) PhoneNumber2 = digits.Substring(first, middle);
                                PhoneNumber3 = digits.Substring(len - last);
                            }
                            else
                            {
                                PhoneNumber1 = digits;
                            }
                        }
                        else
                        {
                            // 小さい桁数は全体をPhoneNumber1へ
                            PhoneNumber1 = digits;
                        }
                    }

                    //Fax番号
                    ScanFaxNumber();

                }
            }
        }

        public void ScanFaxNumber()
        {
            kryptonButton6.Show();
            // Outlookアプリケーションのインスタンスを取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            // 検索したい名前
            string targetName = kryptonListBox3.SelectedItem.ToString();

            // 連絡先を検索
            Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
            Microsoft.Office.Interop.Outlook.ContactItem contact = contactItems.Find($"[FullName] = '{targetName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

            // 例: kryptonListBox3_SelectedIndexChanged 内で contact.BusinessFaxNumber を分割して格納
            FaxNumber1 = FaxNumber2 = FacNumber3 = string.Empty;
            string fax = (contact.BusinessFaxNumber ?? string.Empty).Trim();

            if (!string.IsNullOrEmpty(fax))
            {
                // 数字以外で分割（ハイフンやスペース、括弧などを区切りにする）
                string[] rawParts = System.Text.RegularExpressions.Regex.Split(fax, @"\D+");
                System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
                foreach (var p in rawParts)
                {
                    if (!string.IsNullOrEmpty(p)) parts.Add(p);
                }

                if (parts.Count >= 3)
                {
                    FaxNumber1 = parts[0];
                    FaxNumber2 = parts[1];
                    FacNumber3 = parts[2];
                }
                else if (parts.Count == 2)
                {
                    FaxNumber1 = parts[0];
                    FaxNumber2 = parts[1];
                    FacNumber3 = string.Empty;
                }
                else if (parts.Count == 1)
                {
                    // 桁数に応じて分割（簡易フォールバック）
                    string digits = parts[0];
                    int len = digits.Length;
                    if (len >= 7)
                    {
                        int last = 4;
                        int first = Math.Max(2, len - 7);
                        int middle = len - first - last;
                        if (first > 0 && middle >= 0)
                        {
                            FaxNumber1 = digits.Substring(0, first);
                            if (middle > 0) FaxNumber2 = digits.Substring(first, middle);
                            FacNumber3 = digits.Substring(len - last);
                        }
                        else
                        {
                            FaxNumber1 = digits;
                        }
                    }
                    else
                    {
                        FaxNumber1 = digits;
                    }
                }
            }
        }

        private void kryptonButton16_Click(object sender, EventArgs e)
        {
            ContactsAuth();
        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;
            //名前
            //組織だった場合
            if (kryptonRadioButton4.Checked == true)
            {
                if (Address_NameLabel.Text != "名前を選択してください。")
                {
                    kryptonTextBox3.Text = Address_NameLabel.Text;
                }
                else if (Address_NameLabel.Text != string.Empty)
                {
                    kryptonTextBox3.Text = Address_NameLabel.Text;
                }
            }
            //名前だった場合
            else if (kryptonRadioButton5.Checked == true)
            {
                if (Address_NameLabel.Text != "名前を選択してください。")
                {
                    kryptonTextBox6.Text = Address_NameLabel.Text;
                }
                else if (Address_NameLabel.Text != string.Empty)
                {
                    kryptonTextBox6.Text = Address_NameLabel.Text;
                }
            }

            //会社場所
            if (kryptonCheckBox10.Text != "所在地:〒")
            {
                if (kryptonCheckBox10.Checked == true)
                {
                    kryptonTextBox4.Text = "〒" + loaction;
                }
            }
            //メールアドレス
            if (kryptonCheckBox11.Text != "メールアドレス:")
            {
                if (kryptonCheckBox11.Checked == true)
                {
                    kryptonTextBox7.Text = MailAddress_User;
                    kryptonComboBox8.Text = MailAddress_Domain;
                }
            }
            //電話番号
            if (kryptonCheckBox12.Text != "会社電話番号:")
            {
                if (kryptonCheckBox12.Checked == true)
                {
                    kryptonComboBox6.Text = PhoneNumber1;
                    kryptonTextBox14.Text = PhoneNumber2;
                    kryptonTextBox8.Text = PhoneNumber3;
                }
            }
            //Fax番号
            if (kryptonCheckBox13.Text != "Fax番号:")
            {
                if (kryptonCheckBox13.Checked == true)
                {
                    kryptonComboBox7.Text = FaxNumber1;
                    kryptonTextBox9.Text = FaxNumber2;
                    kryptonTextBox15.Text = FacNumber3;
                }
            }
        }


        async System.Threading.Tasks.Task ContactsAuth()
        {
            this.Enabled = false;

            var availableWindowsHello = await UserConsentVerifier.CheckAvailabilityAsync();
            if (availableWindowsHello != UserConsentVerifierAvailability.Available)
            {
                kryptonButton16.Enabled = false;
                kryptonLabel45.Visible = true;

                this.Enabled = true;

                連絡帳機能をロックToolStripMenuItem.Enabled = false;
                連絡先の追加ToolStripMenuItem.Enabled = false;
                連絡先を削除ToolStripMenuItem.Enabled = false;
                連絡先のリストを更新ToolStripMenuItem.Enabled = false;

                toolStrip2.Enabled = false;
            }
            else
            {
                var result = await UserConsentVerifier.RequestVerificationAsync("Microsoft Outlook の連絡先を取得・使用するには認証してください。");

                if (result == UserConsentVerificationResult.Verified)
                {
                    //認証出来た場合
                    連絡帳機能をロックToolStripMenuItem.Enabled = true;
                    連絡先の追加ToolStripMenuItem.Enabled = true;
                    連絡先を削除ToolStripMenuItem.Enabled = true;
                    連絡先のリストを更新ToolStripMenuItem.Enabled = true;
                    toolStrip2.Enabled = true;

                    kryptonPanel20.Hide();

                    Address_NameLabel.Show();
                    kryptonPanel19.Show();

                    kryptonRibbonGroupButton_AddContact.Enabled = true;
                    kryptonRibbonGroupButton_RemoveContact.Enabled = true;
                    kryptonRibbonGroupButton_UpdateContacts.Enabled = true;
                    kryptonRibbonGroupButton21.Enabled = true;

                    this.Enabled = true;

                    //連絡先取得処理
                    // Outlookアプリケーションを初期化
                    Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                    NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                    // 連絡先フォルダを取得
                    MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                    // 連絡先アイテムを取得
                    Items contactItems = contactsFolder.Items;

                    // 連絡先をループで表示
                    foreach (object item in contactItems)
                    {
                        if (item is ContactItem contact)
                        {
                            kryptonListBox3.Items.Add(contact.FullName);
                        }
                    }
                }
                else
                {
                    //認証をキャンセルした場合
                    kryptonPanel20.Show();

                    連絡先の追加ToolStripMenuItem.Enabled = false;
                    連絡先を削除ToolStripMenuItem.Enabled = false;
                    連絡先のリストを更新ToolStripMenuItem.Enabled = false;
                    連絡帳機能をロックToolStripMenuItem.Enabled = false;
                    toolStrip2.Enabled = false;

                    Address_NameLabel.Hide();
                    kryptonButton6.Hide();
                    kryptonPanel19.Hide();

                    kryptonRibbonGroupButton_AddContact.Enabled = false;
                    kryptonRibbonGroupButton_RemoveContact.Enabled = false;
                    kryptonRibbonGroupButton_UpdateContacts.Enabled = false;
                    kryptonRibbonGroupButton21.Enabled = false;

                    this.Enabled = true;
                }
            }

        }

        private void kryptonCheckBox11_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupButton21_Click(object sender, EventArgs e)
        {
            //認証をキャンセルした場合
            kryptonPanel20.Show();

            Address_NameLabel.Hide();
            kryptonButton6.Hide();
            kryptonPanel19.Hide();

            kryptonRibbonGroupButton_AddContact.Enabled = false;
            kryptonRibbonGroupButton_RemoveContact.Enabled = false;
            kryptonRibbonGroupButton_UpdateContacts.Enabled = false;
            kryptonRibbonGroupButton21.Enabled = false;

            連絡先の追加ToolStripMenuItem.Enabled = false;
            連絡先を削除ToolStripMenuItem.Enabled = false;
            連絡先のリストを更新ToolStripMenuItem.Enabled = false;
            連絡帳機能をロックToolStripMenuItem.Enabled = false;
            toolStrip2.Enabled = false;

            kryptonButton6.Visible = false;
            kryptonListBox3.Items.Clear();

            Address_NameLabel.Text = "名前を選択してください。";
            kryptonCheckBox10.Text = "所在地:";
            kryptonCheckBox11.Text = "メールアドレス:";
            kryptonCheckBox12.Text = "会社電話番号:";
            kryptonCheckBox13.Text = "会社Fax番号:";
        }

        private void kryptonRibbonGroupButton_UpdateContacts_Click(object sender, EventArgs e)
        {
            ContactUpDate();
        }


        public void ContactUpDate()
        {
            kryptonListBox3.Items.Clear();
            //連絡先取得処理
            // Outlookアプリケーションを初期化
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

            // 連絡先アイテムを取得
            Items contactItems = contactsFolder.Items;

            // 連絡先をループで表示
            foreach (object item in contactItems)
            {
                if (item is ContactItem contact)
                {
                    kryptonListBox3.Items.Add(contact.FullName);
                }
            }
        }

        private void kryptonRibbonGroupButton_AddContact_Click(object sender, EventArgs e)
        {
            // Outlookアプリケーションのインスタンス取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();

            // 新規連絡先アイテムを作成
            Microsoft.Office.Interop.Outlook.ContactItem contact =
                outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olContactItem)
                as Microsoft.Office.Interop.Outlook.ContactItem;

            if (contact != null)
            {
                // 追加画面（Outlookの連絡先フォーム）を表示
                contact.Display(true); // true: モーダル表示, false: 非モーダル

            }
            AddContactDialog addContactDialog = new AddContactDialog();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                addContactDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                addContactDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                addContactDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                addContactDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                addContactDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                addContactDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            addContactDialog.ShowDialog();
        }

        private void kryptonRibbonGroupButton_RemoveContact_Click(object sender, EventArgs e)
        {
            DeleteOutlookContactByName(flName);
        }

        public void DeleteOutlookContactByName(string fullName)
        {
            if (kryptonListBox3.SelectedItem != null)
            {
                AddressDeleteWarningTaskDialog addressDeleteWarningTaskDialog = new AddressDeleteWarningTaskDialog();
                //Office2007青色
                if (this.BackColor == Color.FromArgb(191, 219, 255))
                {
                    addressDeleteWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                }
                //Office2007銀色
                else if (this.BackColor == Color.FromArgb(208, 212, 221))
                {
                    addressDeleteWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                }
                //Office2007ブラック
                else if (this.BackColor == Color.FromArgb(83, 83, 83))
                {
                    addressDeleteWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                }
                //Office2010青色
                else if (this.BackColor == Color.FromArgb(187, 206, 230))
                {
                    addressDeleteWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                }
                //Office2010銀色
                else if (this.BackColor == Color.FromArgb(227, 230, 232))
                {
                    addressDeleteWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                }
                //Office2010黒色
                else if (this.BackColor == Color.FromArgb(113, 113, 113))
                {
                    addressDeleteWarningTaskDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                }
                var result = addressDeleteWarningTaskDialog.ShowDialog();
                {
                    if (result == DialogResult.Yes)
                    {
                        fullName = kryptonListBox3.SelectedItem.ToString();
                        // Outlookアプリケーションのインスタンス取得
                        Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                        Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

                        // 連絡先フォルダを取得
                        Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder =
                            outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

                        // 連絡先アイテムを検索
                        Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
                        Microsoft.Office.Interop.Outlook.ContactItem contact =
                            contactItems.Find($"[FullName] = '{fullName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

                        if (contact != null)
                        {
                            contact.Delete(); // 連絡先を削除
                            AddressDeleteDone addressDeleteDone = new AddressDeleteDone();
                            //Office2007青色
                            if (this.BackColor == Color.FromArgb(191, 219, 255))
                            {
                                addressDeleteDone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                            }
                            //Office2007銀色
                            else if (this.BackColor == Color.FromArgb(208, 212, 221))
                            {
                                addressDeleteDone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                            }
                            //Office2007ブラック
                            else if (this.BackColor == Color.FromArgb(83, 83, 83))
                            {
                                addressDeleteDone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                            }
                            //Office2010青色
                            else if (this.BackColor == Color.FromArgb(187, 206, 230))
                            {
                                addressDeleteDone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                            }
                            //Office2010銀色
                            else if (this.BackColor == Color.FromArgb(227, 230, 232))
                            {
                                addressDeleteDone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                            }
                            //Office2010黒色
                            else if (this.BackColor == Color.FromArgb(113, 113, 113))
                            {
                                addressDeleteDone.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                            }
                            addressDeleteDone.ShowDialog();
                            ContactUpDate();
                        }
                        else
                        {
                            AddressDeleteError addressDeleteError = new AddressDeleteError();
                            //Office2007青色
                            if (this.BackColor == Color.FromArgb(191, 219, 255))
                            {
                                addressDeleteError.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                            }
                            //Office2007銀色
                            else if (this.BackColor == Color.FromArgb(208, 212, 221))
                            {
                                addressDeleteError.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                            }
                            //Office2007ブラック
                            else if (this.BackColor == Color.FromArgb(83, 83, 83))
                            {
                                addressDeleteError.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                            }
                            //Office2010青色
                            else if (this.BackColor == Color.FromArgb(187, 206, 230))
                            {
                                addressDeleteError.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                            }
                            //Office2010銀色
                            else if (this.BackColor == Color.FromArgb(227, 230, 232))
                            {
                                addressDeleteError.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                            }
                            //Office2010黒色
                            else if (this.BackColor == Color.FromArgb(113, 113, 113))
                            {
                                addressDeleteError.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                            }
                            addressDeleteError.ShowDialog();
                            ContactUpDate();
                        }
                    }
                }

            }
            else
            {
                SelectAddress selectAddress = new SelectAddress();
                //Office2007青色
                if (this.BackColor == Color.FromArgb(191, 219, 255))
                {
                    selectAddress.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                }
                //Office2007銀色
                else if (this.BackColor == Color.FromArgb(208, 212, 221))
                {
                    selectAddress.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                }
                //Office2007ブラック
                else if (this.BackColor == Color.FromArgb(83, 83, 83))
                {
                    selectAddress.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                }
                //Office2010青色
                else if (this.BackColor == Color.FromArgb(187, 206, 230))
                {
                    selectAddress.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                }
                //Office2010銀色
                else if (this.BackColor == Color.FromArgb(227, 230, 232))
                {
                    selectAddress.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                }
                //Office2010黒色
                else if (this.BackColor == Color.FromArgb(113, 113, 113))
                {
                    selectAddress.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                }
                selectAddress.ShowDialog();
            }

        }

        private void Notepads_kryptonRichTextBox_Notepad_TextChanged_2(object sender, EventArgs e)
        {

        }

        public void AutoSave()
        {
            try
            {
                String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase";

                if (Directory.Exists(str))
                {
                    Notepads_kryptonRichTextBox_Notepad.SaveFile(str + @"\SaveFile.rtf");
                }
                else
                {
                    //フォルダを作成してからファイルを保存
                    Directory.CreateDirectory(str);
                    Notepads_kryptonRichTextBox_Notepad.SaveFile(str + @"\SaveFile.rtf");
                }
            }
            catch { }

        }

        private void kryptonRibbonGroupButton1_NotepadShowExplorer_Click(object sender, EventArgs e)
        {
            String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase";
            if (Directory.Exists(str))
            {
                System.Diagnostics.Process.Start("explorer.exe", str);
            }
        }

        //印刷
        private void kryptonContextMenuItem29_Click(object sender, EventArgs e)
        {
            kryptonLabel1.Text = "出力中...";
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = false;
            Document doc = word.Documents.Add();


            //外枠の余白を設定
            doc.PageSetup.TopMargin = Sheets_TopPanel.Height;
            doc.PageSetup.BottomMargin = Sheets_ButtomPanel.Height;
            doc.PageSetup.LeftMargin = Sheets_LeftPanel.Width;
            doc.PageSetup.RightMargin = Sheets_RightPanel.Width;

            foreach (Range range in doc.StoryRanges)
            {
                range.Font.Size = 10; // フォントサイズを10に設定
            }

            //発行番号
            if (Sheets_NumberLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph1 = doc.Paragraphs.Add();
                paragraph1.Range.Text = Sheets_NumberLabel.Text;
                paragraph1.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph1.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph1.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph1.Range.InsertParagraphAfter();
            }
            //日付
            if (Sheets_DateLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph2 = doc.Paragraphs.Add();
                paragraph2.Range.Text = Sheets_DateLabel.Text;
                paragraph2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph2.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph2.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph2.Range.InsertParagraphAfter();
            }
            //相手先会社名
            if (Sheets_AddressCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph3 = doc.Paragraphs.Add();
                paragraph3.Range.Text = Sheets_AddressCompanyLabel.Text;
                paragraph3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph3.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph3.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph3.Range.InsertParagraphAfter();
            }
            //相手先氏名
            if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph4 = doc.Paragraphs.Add();
                paragraph4.Range.Text = Sheets_AddressTitleAndNameLabel.Text;
                paragraph4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph4.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph4.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph4.Range.InsertParagraphAfter();
            }
            //発信者会社名
            if (Sheets_CallerCompanyLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph5 = doc.Paragraphs.Add();
                paragraph5.Range.Text = Sheets_CallerCompanyLabel.Text;
                paragraph5.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph5.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph5.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph5.Range.InsertParagraphAfter();
            }
            //発信者所在地
            if (Sheets_CallerLocationLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph6 = doc.Paragraphs.Add();
                paragraph6.Range.Text = Sheets_CallerLocationLabel.Text;
                paragraph6.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph6.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph6.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph6.Range.InsertParagraphAfter();
            }
            //発信者建物名と階数
            if (Sheets_BuildingNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph7 = doc.Paragraphs.Add();
                paragraph7.Range.Text = Sheets_BuildingNameLabel.Text;
                paragraph7.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph7.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph7.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph7.Range.InsertParagraphAfter();
            }
            //発信者氏名
            if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph8 = doc.Paragraphs.Add();
                paragraph8.Range.Text = Sheets_CallerTitleAndNameLabel.Text;
                paragraph8.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph8.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph8.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph8.Range.InsertParagraphAfter();
            }
            //メールアドレス
            if (Sheets_CallerMallAddressLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph9 = doc.Paragraphs.Add();
                paragraph9.Range.Text = "メールアドレス:" + Sheets_CallerMallAddressLabel.Text;
                paragraph9.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph9.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph9.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph9.Range.InsertParagraphAfter();
            }
            //電話番号
            if (Sheets_CallerTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph10 = doc.Paragraphs.Add();
                paragraph10.Range.Text = "電話番号:" + Sheets_CallerTelLabel.Text;
                paragraph10.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph10.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph10.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph10.Range.InsertParagraphAfter();
            }
            //Fax番号
            if (Sheets_CallerFaxTelLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph11 = doc.Paragraphs.Add();
                paragraph11.Range.Text = "Fax番号:" + Sheets_CallerFaxTelLabel.Text;
                paragraph11.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph11.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph11.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph11.Range.InsertParagraphAfter();
            }
            //表題
            if (Sheets_TitleButton.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph12 = doc.Paragraphs.Add();
                if (kryptonRibbonButton_Bold.Checked == true)
                {
                    paragraph12.Range.Bold = 1;
                }
                if (kryptonRibbonButton_Italic.Checked == true)
                {
                    paragraph12.Range.Italic = 1;
                }
                if (kryptonContextMenuItem15.Checked == true)
                {
                    paragraph12.Range.Underline = WdUnderline.wdUnderlineSingle;
                }
                if (kryptonContextMenuItem16.Checked == true)
                {
                    paragraph12.Range.Font.StrikeThrough = 1;
                }
                paragraph12.Range.Font.Size = (int)kryptonTextBox10.StateCommon.Content.Font.Size;
                paragraph12.Range.Text = Sheets_TitleButton.Text;
                paragraph12.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                SetWordRangeColor(paragraph12.Range, Sheets_TitleButton.ForeColor);
                paragraph12.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph12.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph12.Range.InsertParagraphAfter();

            }
            //あいさつ文
            if (Sheets_ContentLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph13 = doc.Paragraphs.Add();
                paragraph13.Range.Bold = 0;
                paragraph13.Range.Italic = 0;
                paragraph13.Range.Underline = WdUnderline.wdUnderlineNone;
                paragraph13.Range.Font.StrikeThrough = 0;
                paragraph13.Range.Font.Size = 10;
                paragraph13.Range.Text = Sheets_ContentLabel.Text;
                paragraph13.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph13.Range.Font.Color = WdColor.wdColorBlack;
                paragraph13.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph13.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph13.Range.InsertParagraphAfter();
            }
            //内容
            try
            {
                int LinesCount = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox12.Lines[LinesCount];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount = LinesCount + 1;
                    if (LinesCount == kryptonTextBox12.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //結語
            //kryptonRibbonGroupCheckBoxにチェックがない場合に動作する
            if (kryptonRibbonGroupCheckBox1.Checked != true)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph14 = doc.Paragraphs.Add();
                paragraph14.Range.Text = Sheet_ConclusionLabel.Text;
                paragraph14.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph14.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph14.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph14.Range.InsertParagraphAfter();
            }
            kryptonLabel1.Text = "出力完了";
            stausUpdate();
            //記
            if (Sheet_NoteLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph15 = doc.Paragraphs.Add();
                paragraph15.Range.Text = Sheet_NoteLabel.Text;
                paragraph15.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph15.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph15.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph15.Range.InsertParagraphAfter();
            }
            //記し書き
            try
            {
                int LinesCount2 = 0;
                while (true)
                {
                    Microsoft.Office.Interop.Word.Paragraph W_Contents = doc.Paragraphs.Add();
                    W_Contents.Range.Text = kryptonTextBox13.Lines[LinesCount2];
                    W_Contents.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    W_Contents.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                    W_Contents.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                    W_Contents.Range.InsertParagraphAfter();
                    LinesCount2 = LinesCount2 + 1;
                    if (LinesCount2 == kryptonTextBox13.Lines.Length)
                    {
                        break;
                    }
                }
            }
            catch { }
            //以上
            if (Sheets_EndLabel.Text != string.Empty)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph16 = doc.Paragraphs.Add();
                paragraph16.Range.Text = Sheets_EndLabel.Text;
                paragraph16.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                paragraph16.Format.SpaceAfter = 0; // 段落後の間隔（ポイント単位）
                paragraph16.Format.SpaceBefore = 0; // 段落前の間隔（ポイント単位）
                paragraph16.Range.InsertParagraphAfter();
            }
            GC.Collect();


            PrintDialog pd = new PrintDialog();
            pd.UseEXDialog = true;
            if (pd.ShowDialog() == DialogResult.OK)
            {
                word.ActivePrinter = pd.PrinterSettings.PrinterName;
                doc.PrintOut();
            }

            try
            {
                doc.Close(false);
                word.Quit();
            }
            catch { }
        }

        private void kryptonButton17_Click(object sender, EventArgs e)
        {
            try
            {
                if (kryptonComboBox18.Text == "発行元部署")
                {
                    string str = kryptonTextBox11.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox11.Text = str;
                }
                else if (kryptonComboBox18.Text == "宛先の組織・会社名")
                {
                    string str = kryptonTextBox1.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox1.Text = str;
                }
                else if (kryptonComboBox18.Text == "宛先の肩書き")
                {
                    string str = kryptonComboBox10.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox10.Text = str;
                }
                else if (kryptonComboBox18.Text == "宛先の氏名")
                {
                    string str = kryptonTextBox2.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox2.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の会社名")
                {
                    string str = kryptonTextBox3.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox3.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の所在地")
                {
                    string str = kryptonTextBox4.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox4.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の建物名")
                {
                    string str = kryptonTextBox5.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox5.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の肩書き")
                {
                    string str = kryptonComboBox9.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox9.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の氏名")
                {
                    string str = kryptonTextBox6.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox6.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者のメールアドレス(ユーザー)")
                {
                    string str = kryptonTextBox7.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox7.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者のメールアドレス(ドメイン)")
                {
                    string str = kryptonComboBox8.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox8.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の電話番号(1)")
                {
                    string str = kryptonComboBox6.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox6.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の電話番号(2)")
                {
                    string str = kryptonTextBox14.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox14.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者の電話番号(3)")
                {
                    string str = kryptonTextBox8.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox8.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者のFax番号(1)")
                {
                    string str = kryptonComboBox7.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox7.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者のFax番号(2)")
                {
                    string str = kryptonTextBox9.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox9.Text = str;
                }
                else if (kryptonComboBox18.Text == "発信者のFax番号(3)")
                {
                    string str = kryptonTextBox15.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox15.Text = str;
                }
                else if (kryptonComboBox18.Text == "表題")
                {
                    string str = kryptonTextBox10.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonTextBox10.Text = str;
                }
                else if (kryptonComboBox18.Text == "頭語")
                {
                    string str = kryptonComboBox2.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox2.Text = str;
                }
                else if (kryptonComboBox18.Text == "候文")
                {
                    string str = kryptonComboBox11.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox11.Text = str;
                }
                else if (kryptonComboBox18.Text == "前文")
                {
                    string str = kryptonComboBox3.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox3.Text = str;
                }
                else if (kryptonComboBox18.Text == "感謝のあいさつ")
                {
                    string str = kryptonComboBox4.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox4.Text = str;
                }
                else if (kryptonComboBox18.Text == "結語")
                {
                    string str = kryptonComboBox5.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox5.Text = str;
                }
                else if (kryptonComboBox18.Text == "内容")
                {
                    string str = kryptonTextBox12.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox12.Text = str;
                }
                else if (kryptonComboBox18.Text == "記し書き")
                {
                    string str = kryptonTextBox13.Text.Replace(kryptonTextBox30.Text, kryptonTextBox31.Text);
                    kryptonComboBox13.Text = str;
                }
            }
            catch { }

        }

        private void kryptonRibbonGroupButton16_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupButton16.Checked == true)
            {

                Transition
                    .With(kryptonPanel21, nameof(Height), 36)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = true;
                置換ToolStripMenuItem.Checked = true;

                kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.SlantOutsizeFar;
            }
            else if (kryptonRibbonGroupButton16.Checked == false)
            {
                kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.OneNote;
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = false;
                置換ToolStripMenuItem.Checked = false;
            }

        }

        private void kryptonRibbonGroupButton22_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(System.Windows.Forms.Application.ExecutablePath);
        }

        private void kryptonButton18_Click(object sender, EventArgs e)
        {
            kryptonTextBox16.Text = String.Empty;
            kryptonTextBox17.Text = String.Empty;
            kryptonComboBox12.Text = String.Empty;
            kryptonTextBox18.Text = String.Empty;
            kryptonTextBox19.Text = String.Empty;
            kryptonTextBox32.Text = String.Empty;
            kryptonTextBox20.Text = String.Empty;
            kryptonNumericUpDown3.Value = 1;
            kryptonComboBox13.Text = String.Empty;
            kryptonTextBox21.Text = String.Empty;
            kryptonTextBox22.Text = String.Empty;
            kryptonComboBox14.Text = String.Empty;
            kryptonComboBox15.Text = String.Empty;
            kryptonTextBox23.Text = String.Empty;
            kryptonTextBox24.Text = String.Empty;
            kryptonComboBox16.Text = String.Empty;
            kryptonTextBox26.Text = String.Empty;
            kryptonTextBox25.Text = String.Empty;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            splashWindow.Close();
            splashWindow.Dispose();

        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            kryptonNumericUpDown4.Value = Properties.Settings.Default.Space_Top;
            kryptonNumericUpDown7.Value = Properties.Settings.Default.Space_Buttom;
            kryptonNumericUpDown5.Value = Properties.Settings.Default.Space_Left;
            kryptonNumericUpDown6.Value = Properties.Settings.Default.Space_Right;
        }

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

        private void kryptonLinkLabel1_LinkClicked(object sender, EventArgs e)
        {
            Properties.Settings.Default.ShowNotepadWarningPanel = false;
            Properties.Settings.Default.Save();

            Transition
                .With(WarningPanel1, nameof(Height), 0)
                .CriticalDamp(TimeSpan.FromSeconds(0.4));
        }

        private void kryptonRibbonGroupButton_TextReset_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.ShowResetDialog == true)
            {

                TextResetDialog textResetDialog = new TextResetDialog();
                //Office2007青色
                if (this.BackColor == Color.FromArgb(191, 219, 255))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                }
                //Office2007銀色
                else if (this.BackColor == Color.FromArgb(208, 212, 221))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                }
                //Office2007ブラック
                else if (this.BackColor == Color.FromArgb(83, 83, 83))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                }
                //Office2010青色
                else if (this.BackColor == Color.FromArgb(187, 206, 230))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                }
                //Office2010銀色
                else if (this.BackColor == Color.FromArgb(227, 230, 232))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                }
                //Office2010黒色
                else if (this.BackColor == Color.FromArgb(113, 113, 113))
                {
                    textResetDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                }
                textResetDialog.ShowDialog();
                if (textResetDialog.DialogResult == DialogResult.Yes)
                {
                    SetSheetSpace();
                    SetSheetText();
                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    kryptonNumericUpDown1.Value = 0;
                    kryptonDateTimePicker1.Value = DateTime.Today;

                    kryptonCheckBox3.Checked = false;
                    kryptonCheckBox2.Checked = false;

                    if (Properties.Settings.Default.IsUseEraName == true)
                    {
                        kryptonCheckBox1.Checked = true;
                    }
                    else
                    {
                        kryptonCheckBox1.Checked = false;
                    }
                }
            }
            else
            {
                SetSheetSpace();
                SetSheetText();
                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                kryptonNumericUpDown1.Value = 0;
                kryptonDateTimePicker1.Value = DateTime.Today;

                kryptonCheckBox3.Checked = false;
                kryptonCheckBox2.Checked = false;

                if (Properties.Settings.Default.IsUseEraName == true)
                {
                    kryptonCheckBox1.Checked = true;
                }
                else
                {
                    kryptonCheckBox1.Checked = false;
                }
            }
        }

        //上
        private void kryptonContextMenuItem22_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = Properties.Settings.Default.Space_Top;

            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = Properties.Settings.Default.Space_Top;
        }

        //下
        private void kryptonContextMenuItem23_Click(object sender, EventArgs e)
        {
            Sheets_ButtomPanel.Height = Properties.Settings.Default.Space_Buttom;

            kryptonRibbonGroupNumericUpDown1.Value = Properties.Settings.Default.Space_Buttom;
        }

        //右
        private void kryptonContextMenuItem24_Click(object sender, EventArgs e)
        {
            Sheets_RightPanel.Width = Properties.Settings.Default.Space_Right;

            kryptonRibbonGroupNumericUpDown2.Value = Properties.Settings.Default.Space_Right;
        }

        //左
        private void kryptonContextMenuItem25_Click(object sender, EventArgs e)
        {
            Sheets_LeftPanel.Width = Properties.Settings.Default.Space_Left;

            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = Properties.Settings.Default.Space_Left;
        }


        //すべて
        private void kryptonContextMenuItem26_Click(object sender, EventArgs e)
        {
            Sheets_TopPanel.Height = Properties.Settings.Default.Space_Top;
            Sheets_ButtomPanel.Height = Properties.Settings.Default.Space_Buttom;
            Sheets_LeftPanel.Width = Properties.Settings.Default.Space_Left;
            Sheets_RightPanel.Width = Properties.Settings.Default.Space_Right;

            kryptonRibbonGroupNumericUpDown_VerticalSpace.Value = Properties.Settings.Default.Space_Top;
            kryptonRibbonGroupNumericUpDown1.Value = Properties.Settings.Default.Space_Buttom;
            kryptonRibbonGroupNumericUpDown_WidthSpace.Value = Properties.Settings.Default.Space_Left;
            kryptonRibbonGroupNumericUpDown2.Value = Properties.Settings.Default.Space_Right;
        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            kryptonPanel21.Show();

            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                menuStripPanel.Hide();
                kryptonRibbon.Show();
                this.AllowFormChrome = true;
            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                menuStripPanel.Show();
                kryptonRibbon.Hide();
                this.AllowFormChrome = false;
            }



            kryptonPanel5.Show();

            //設定の復元
            SetSettings();
            if (kryptonRibbonGroupButton16.Checked == true)
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 36)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = true;
            }
            else
            {
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                kryptonRibbonGroupButton16.Checked = false;
            }

            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }


            kryptonPage2.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;

            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuEase Designer";
            }
            else
            {
                this.Text = "無題 - DocuEase Designer";
            }

            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;

        }

        private void kryptonRibbonGroupButton_NotepadSave_Click(object sender, EventArgs e)
        {
            AutoSave();
        }

        //TreeView1を選択したとき
        private void treeView1_AfterSelect_1(object sender, TreeViewEventArgs e)
        {
            if (treeView1.SelectedNode == ultraMiniNode1)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "注文書";
                kryptonLabel35.Text = "商品やサービス等を注文する側が、注文を受ける側に対して使用します。";

                kryptonTextBox27.Text = "　さて、このたびはお見積書をご送付いただきありがとうございます。\r\nつきましては、下記のとおりご注文申し上げますので、よろしくお願い申し上げます。";
                kryptonTextBox28.Text = "商品名称:\r\n商品番号:\r\n数量:\r\n単価:\r\n値段:\r\n\r\n小計:\r\n割引合計:\r\n税金:\r\n合計:\r\n\r\n備考:";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode2)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "承諾書";
                kryptonLabel35.Text = "宛先に対して約款等を遵守させるときに使用します。";

                kryptonTextBox27.Text = "私は、○○について、下記のとおり遵守し同意します。";
                kryptonTextBox28.Text = "1.\r\n2.\r\n3.";

                kryptonPanel17.Show();
            }
            else if (treeView1.SelectedNode == ultraMiniNode3)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "依頼文（○○のお願い）";
                kryptonLabel35.Text = "宛先に協力や依頼事をお願いする文書です。";

                kryptonTextBox27.Text = "　さて、突然のお願いで恐縮ですが、現在進行中の○○に関して、貴社のご協力をお願いしたくご連絡いたしました。\r\n具体的には、○○の件についてご意見をいただけますと幸いです。お忙しいところ大変恐縮ですが、何卒よろしくお願い申し上げます。";
                kryptonTextBox28.Text = string.Empty;

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode4)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "照会文（○○のご照会のお願い）";
                kryptonLabel35.Text = "宛先に対して事務上の疑問点や不明点を問い合わせを行う文書です。";

                kryptonTextBox27.Text = "　さて、○○について事務上の参考にさせていただきたいので、下記の事項について○月○日までにご回答くださりますようお願い申し上げます。";
                kryptonTextBox28.Text = "1.\r\n2.\r\n3.";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode5)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "回答文（○○に対するお問い合わせの件(回答)）";
                kryptonLabel35.Text = "宛先の問い合わせに対して回答する文書です。";

                kryptonTextBox27.Text = "　さて、このたびは○○の件にお問い合わせいただき誠にありがとうございます。\r\n　つきましては、下記のとおりご回答を申し上げます。\r\n(回答の内容)\r\n　なお、ご不明な点がございましたら下記担当までお問い合わせください。\r\n\r\nまずは、書面をもちましてご回答申し上げます。今後とも変わらずお引き立てのほどよろしくお願い申し上げます。";
                kryptonTextBox28.Text = "・お問い合わせ先　　○○部　03(0000)0000　担当○○まで";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode6)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "催促文（○○についての○○のお願い）";
                kryptonLabel35.Text = "宛先に対してお支払いや提出・返却などを催促する文書です。";

                kryptonTextBox27.Text = "　さて、令和○年○月○日にて弊社より○○しました○○が、○○予定日の令和○年○月○日を過ぎた本日になってもいまだ○○いただいておりません。\r\nつきましては、至急下記のとおりまで○○くださいますようお願い申し上げます。\r\nまずは、書面をもちましてご通知申し上げます。\r\n　なお、本状と行き違いにより○○いただいた場合は、悪しからずご容赦ください。";
                kryptonTextBox28.Text = "1.(物品名または金額)\r\n2.(期限)\r\n3.(方法)\r\n4.(問い合わせ先)";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode7)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "断り文（○○ご辞退の件）";
                kryptonLabel35.Text = "宛先に対して辞退を伝えるための文書です。";

                kryptonTextBox27.Text = "　さて、このたびは○○を○○していただき誠にありがとうございました。\r\n早速貴社のご提案を社内で慎重に検討しましたが、○○のため、誠に勝手ながら今回はご辞退申し上げます。ご要望にお応えできなくなってしまい誠に申し訳ございませんでした。\r\n　(辞退した簡潔な理由　例:貴社の提案をご辞退いたしました理由としまして、貴社ご希望の条件では弊社ではお受けできかねるためです。)\r\n　なにとぞ諸事情をお汲み取りのうえ、ご了承くださいますようお願い申し上げます。\r\n　つきましては、略儀ながら書面をもちまして○○の辞退のお知らせを申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode8)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "交渉文（○○のお願い）";
                kryptonLabel35.Text = "宛先に価格等の交渉を行うための文書です。";

                kryptonTextBox27.Text = "　さて、現在御社から仕入れております「○○」について、○○の○○をお願いしたく、ご連絡させていただきました。\r\n　○○などにより、思うような販売の成果が得られず、苦戦を強いてられているため御社にご協力を賜りたく存じます\r\nつきましては大変厚かましく勝手なお願いで恐縮ですが○○を○○％ほど○○していただけないでしょうか?\r\n　取り急ぎ書面にてお願い申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode9)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "抗議文";
                kryptonLabel35.Text = "発信者が宛先に対して抗議の意を表すためも文書です。";

                kryptonTextBox27.Text = "　令和○年○月○日午後〇時〇分にて、お客様が店内で○○行為を行ったことについて、容疑者対にしここに厳重な抗議をいたします。\r\n　店内で○○行為は、弊社では決して容認するものではなく常識の範囲内をはるかに越え、犯罪行為に匹敵するものです。\r\n　つきましては、当行為で逮捕された容疑者に対し、エリアマネージャーなどによる事情説明と謝罪を強く求めるものとします。\r\n　なお、事情説明の内容によって法的措置をとることも検討しております。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode10)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "お詫び文（○○についてのお詫び）";
                kryptonLabel35.Text = "発信者に対して謝罪を意を表すための文書です。";

                kryptonTextBox27.Text = "　さて、令和○年○月○日に発売いたしました「○○」につきまして製品上の欠陥が見つかったとご指摘をいただいたことに消費者の方々や関係企業などに対してご迷惑をおかけして誠に申し訳なく、深くお詫び申し上げます。\r\n　再度社内で当該製品を確認いたしましたところ、○○の部分が破損していることがわかりました。\r\n　現在、製品の無償返却や製品の問い合わせを受け付けておりますので何卒、よろしくお願い申し上げます。\r\n　今後、このような失態が起きないよう、社内では製品の検査体制や社内規則を徹底的に強化いたしますので、どうか今後とも変わらぬお引き立てをお願い申し上げます。\r\n　まずは、書面をもちまして再度、心よりお詫び申し上げます。";
                kryptonTextBox28.Text = "問い合わせ先: 東京都渋谷区渋谷○○番地○○号　○○ビル ○○階 03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode11)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "取り消し文（○○の注文の取り消し）";
                kryptonLabel35.Text = "発信者に対して取り消しをお願いするための文書です。";

                kryptonTextBox27.Text = "　さて、令和○年○月○日にご注文申し上げた○○について製品上の欠陥が見つかったため、ご迷惑をおかけし誠に申し訳ございませんが、当該製品の注文を取り消しをここに通知します。\r\n　今回の件につきましては、製品の○○の部分が破損していることを社内で発覚し、製品の注文・発送中止等をいたしました次第です。\r\n　お客様のご要望にお応えできなくなってしまい大変深くお詫び申し上げます。\r\n　まずは、略儀ながら書面をもちまして、注文中止の件につきまして、重ねてお詫び申し上げます。";
                kryptonTextBox28.Text = "問い合わせ先: 東京都渋谷区渋谷○○番地○○号　○○ビル ○○階 03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode12)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "あいさつ文（○○新代表取締役社長の就任のお祝いのご挨拶）";
                kryptonLabel35.Text = "発信者に対してお祝いやあいさつの言葉を残す文書です。";

                kryptonTextBox27.Text = "　さて、貴社にてご就任されました○○新代表取締役社長につきまして、社員一同、大変誠に嬉しく思い、心よりお祝い申し上げます。\r\n　かねてより、弊社と提携を○○事業を進めておりましたが、このたび、○○を令和○年○月○日に発売することとなり、新しい未来を創造する日に一歩前進いたしました。弊社ではお客様がより良い生活体験が享受できますよう貴社と緊密な連携を図ることをお約束します。\r\n　今後も貴社がますます大きくご繫栄されることを切にお祈り申し上げます。\r\n　略儀ながら書中をもちましてご挨拶申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode13)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "お祝い文（○○様のご結婚のお祝い）";
                kryptonLabel35.Text = "発信者に対してお祝いの言葉を残す文書です。";

                kryptonTextBox27.Text = "　このたびは、お二人のご結婚、誠におめでとうございます。\r\n   お二人の新生活の門出を心よりお祝い申し上げます。\r\n　これから二人三脚ですばらしい家庭を築かれることを切にお祈り申し上げます。\r\n　ほんのささやかなではございますが、お祝いの品を送らせていただきました。\r\n　お二人の末永い幸せを心よりご期待しております。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode14)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "招待文（○○の新製品発表会のご案内）";
                kryptonLabel35.Text = "発信者に対してイベントの招待を行うための文書です。";

                kryptonTextBox27.Text = "　さて、かねてより開発を重ねてまいりました「○○」がを発売する運びとなりました。\r\n　○○は従来の商品よりも○％向上しており効果の向上を期待できます。さらに○○には「○○」機能を備えており使用することでより簡単に○○の時間を省くことが可能となります。\r\n　なお、この場をお借りして、ささやかながら、当該製品に対するご意見をお伺いし、今後の技術向上に役立てさせていただきたいと存じます。\r\n　ご多忙中、恐れ入りますが、ぜひご出席賜りますようお願い申し上げます。";
                kryptonTextBox28.Text = "1.日時 ○月○日　午後13時\r\n2.場所　東京都港区高輪○○番地○号　プレックス○○　第1ホール\r\n3.お問い合わせ先　東京都渋谷区渋谷○○番地○○号　○○ビル ○○階 03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode15)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "お礼文（○○のお礼）";
                kryptonLabel35.Text = "発信者にお礼を行うための文書です。";

                kryptonTextBox27.Text = "　さて、このたびは、ご多忙中にもかかわらず、○○していただき誠にありがとうございます。\r\n　おかげをもちまして、○○を無事に成功のうちに終えることができました。これもひとえに○○様のご尽力で成功を納めることができ改めて深く感謝いたしております。\r\n　今後とも、ますますの末永いご活躍を社員一同切にお祈り申し上げます。\r\n　まずは略儀にてお礼申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode16)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "案内文（○○のご案内）";
                kryptonLabel35.Text = "発信者に対してイベントや書類の送付などの案内を行うための文書です";

                kryptonTextBox27.Text = "　さて、弊社の事業内容をより深くご理解いただくために○○を下記のとおり開催いたしますのでご案内申し上げます。今回、○○の啓発活動や事業内容について分かりやすく解説するとともに弊社での貢献活動による実績についても発表を行いたいと思います。\r\n　つきましては、ご多忙の中恐縮ですが、万障繰り合わせの上是非ともご参加賜りますようお願い申し上げます。\r\n  略式ながら書面にてご案内申し上げます。";
                kryptonTextBox28.Text = "1.日時　○○○○年○月○日　(月)　○時より\n2.場所　午後13時\r\n2.場所　東京都港区高輪○○番地○号　プレックス○○　第1ホール\r\n3.参加料金　○○○○円\r\n4.参加方法　当日、第１ホールのエントランスホール内に常駐しております受付スタッフにお申し付けください。\r\n4.お問い合わせ先　企画部○○課　担当　○○○○まで\r\n　　電話　03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode17)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "通知文（○○の通知）";
                kryptonLabel35.Text = "発信者に対してお知らせを行う文書です。";

                kryptonTextBox27.Text = "　さて、このたび弊社では、○○(概要)につきまして、下記のとおり（開催、実施、変更、）いたしますのでここにご通知申し上げます。\r\n　（詳細な内容）\r\n　なお、ご不明な点がございましたら、下記担当者までお問い合わせください。";
                kryptonTextBox28.Text = "1.日時　○○○○年○月○日　(月)　○時より\n2.場所　午後13時\r\n2.場所　東京都港区高輪○○番地○号　プレックス○○　第1ホール\r\n3.お問い合わせ先　企画部○○課　担当　○○○○まで\r\n　　電話　03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode18)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "年賀文（あけましておめでとうございます）";
                kryptonLabel35.Text = "発信者に対して新年のお知らせと社員の努力表明を行う文書です。";

                kryptonTextBox27.Text = "　さて、突然ですが、旧年中はひとかたならぬお引き立てを与りまして、厚く御礼申し上げます。\r\n　旧年では○○の発売により御社にとってめざましい功績を収めることができましたが、本年も御社の成長に少しでも貢献できますよう、社員一同が一丸となり全身全霊で油断せず成果上げてゆくことをお約束します。\r\n　本年も倍旧のお引き立てのほど切にお願い申し上げます。\r\n　まずは、新年のご挨拶と社員一同の努力表明の書面とさせていただきます。\r\n　";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode19)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "季節の挨拶文（ご挨拶）";
                kryptonLabel35.Text = "発信者に対して季節の挨拶を表す文書です。";

                kryptonTextBox27.Text = "　さて、ささやかではございますが、季節のご挨拶と感謝と致します。\r\n　日頃の感謝として心ばかりの粗品をお送り申し上げますので、今後ともご支援とご厚情を賜りますようお願い申し上げます。\r\n　これからの季節、寒暖差が激しい時期でありますので、貴社の皆様方におかれましては、ご健康とご活躍をお祈り申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Hide();
            }
            else if (treeView1.SelectedNode == ultraMiniNode20)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "見舞い文（ご挨拶）";
                kryptonLabel35.Text = "発信者に対して病気や怪我などを気遣う文書です。";

                kryptonTextBox27.Text = "　昨日、弊社の社員が○○様が転倒し病院に搬送された旨を伝え聞きました。弊社社員一同、大変驚きを隠せず、ご心配申し上げております。\r\n　知らなかったとはいえ、お見舞いが遅れてしまい大変申し訳ございません。\r\n　幸い、術後の経過は良好のことですが、ご家族の皆様には、さぞやご心配のことでしょう。\r\n　看病のお疲れが出ませんように、どうかご自愛ください。\r\n　一日でも早く、○○様がお元気でいられますよう社員一同心よりお祈り申し上げます。\r\n　近いうちに病院に向かいたいと存じますが、まずは取り急ぎお見舞い申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Show();
            }
            else if (treeView1.SelectedNode == hyperTreeNode1)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "個人宛見舞い文（暑中お見舞い申し上げます）";
                kryptonLabel35.Text = "家族や親戚などに挨拶を伝えるための文書です。";

                kryptonTextBox27.Text = "　暑さ厳しき日がつづいておりますがお変わりございませんか。私たちもおかげをもちまして元気に過ごしております。\r\n　お身体に気を付けて存分に夏をお楽しみください。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Show();
            }
            else if (treeView1.SelectedNode == ultraMiniNode21)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "お悔やみ文（ご遺族の方へ）";
                kryptonLabel35.Text = "ご遺族にお悔やみを伝えるための文書です。";

                kryptonTextBox27.Text = "　○○様のご訃報のに接し、謹んでお悔やみを申し上げます。\r\n　社員一同驚きを隠せず、残念でありません。また本来であればご葬儀に参列すべきところですが、事情によりかなわず、誠に申し訳ございません。\r\n　心ばかりではありますが、ご香典を同封しておりますので、ご霊前にお供えくださりますようお願い申し上げます。\r\n　○○様の安らかなご冥福をお祈り申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Show();
            }

        }

        //TreeView2を選択したとき
        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeView2.SelectedNode == miniTreeNode22)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();
                kryptonLabel34.Text = "通達文（○○に関する○○のお願い）";
                kryptonLabel35.Text = "宛先に対して通達と指示を行う文書です。";

                kryptonTextBox27.Text = "　このたびは、○○することにあたって、下記のとおり実施していただきますようお願い申し上げます。";
                kryptonTextBox28.Text = "1.\r\n2.\r\n3.";

                kryptonPanel17.Show();
            }
            if (treeView2.SelectedNode == miniTreeNode23)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();
                kryptonLabel34.Text = "指示文（○○に関する○○のお願い）";
                kryptonLabel35.Text = "宛先に対して通達と指示を行う文書です。(通達文を参照)";

                kryptonTextBox27.Text = "　このたびは、○○することにあたって、下記のとおり実施していただきますようお願い申し上げます。";
                kryptonTextBox28.Text = "1.\r\n2.\r\n3.";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode24)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();
                kryptonLabel34.Text = "依頼文（○○のお願い）";
                kryptonLabel35.Text = "宛先に協力や依頼事をお願いする文書です。(「社外文書」セクションの「依頼文」を参照)";

                kryptonTextBox27.Text = "　さて、突然のお願いで恐縮ですが、現在進行中の○○に関して、貴社のご協力をお願いしたくご連絡いたしました。\r\n具体的には、○○の件についてご意見をいただけますと幸いです。お忙しいところ大変恐縮ですが、何卒よろしくお願い申し上げます。";
                kryptonTextBox28.Text = string.Empty;

                kryptonTextBox28.Text = "1.\r\n2.\r\n3.";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode25)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "照会文（○○のご照会のお願い）";
                kryptonLabel35.Text = "宛先に対して事務上の疑問点や不明点を問い合わせを行う文書です。（「社外文書」セクションの「照会文」を参照）";

                kryptonTextBox27.Text = "　さて、○○について事務上の参考にさせていただきたいので、下記の事項について○月○日までにご回答くださりますようお願い申し上げます。";
                kryptonTextBox28.Text = "1.\r\n2.\r\n3.";

                kryptonPanel17.Hide();
            }
            else if (treeView2.SelectedNode == miniTreeNode26)
            {

                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "回答文（○○に対するお問い合わせの件(回答)）";
                kryptonLabel35.Text = "宛先の問い合わせに対して回答する文書です。（「社外文書」セクションの「回答文」を参照）";

                kryptonTextBox27.Text = "　さて、このたびは○○の件にお問い合わせいただき誠にありがとうございます。\r\n　つきましては、下記のとおりご回答を申し上げます。\r\n(回答の内容)\r\n　なお、ご不明な点がございましたら下記担当までお問い合わせください。\r\n\r\nまずは、書面をもちましてご回答申し上げます。今後とも変わらずお引き立てのほどよろしくお願い申し上げます。";
                kryptonTextBox28.Text = "・お問い合わせ先　　○○部　03(0000)0000　担当○○まで";

                kryptonPanel17.Hide();
            }
            else if (treeView2.SelectedNode == miniTreeNode27)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "通知文（○○の通知）";
                kryptonLabel35.Text = "宛先に対してお知らせを行う文書です。（「社外文書」セクションの「通知文」を参照）";

                kryptonTextBox27.Text = "　さて、このたび弊社では、○○(概要)につきまして、下記のとおり（開催、実施、変更、）いたしますのでここにご通知申し上げます。\r\n　（詳細な内容）\r\n　なお、ご不明な点がございましたら、下記担当者までお問い合わせください。";
                kryptonTextBox28.Text = "1.日時　○○○○年○月○日　(月)　○時より\n2.場所　午後13時\r\n2.場所　東京都港区高輪○○番地○号　プレックス○○　第1ホール\r\n3.お問い合わせ先　企画部○○課　担当　○○○○まで\r\n　　電話　03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView2.SelectedNode == miniTreeNode28)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "案内文（○○のご案内）";
                kryptonLabel35.Text = "宛先に対してイベントや書類の送付などの案内を行うための文書です。（「社外文書」セクションの「案内文」を参照）";

                kryptonTextBox27.Text = "　さて、弊社の事業内容をより深くご理解いただくために○○を下記のとおり開催いたしますのでご案内申し上げます。今回、○○の啓発活動や事業内容について分かりやすく解説するとともに弊社での貢献活動による実績についても発表を行いたいと思います。\r\n　つきましては、ご多忙の中恐縮ですが、万障繰り合わせの上是非ともご参加賜りますようお願い申し上げます。\r\n  略式ながら書面にてご案内申し上げます。";
                kryptonTextBox28.Text = "1.日時　○○○○年○月○日　(月)　○時より\n2.場所　午後13時\r\n2.場所　東京都港区高輪○○番地○号　プレックス○○　第1ホール\r\n3.参加料金　○○○○円\r\n4.参加方法　当日、第１ホールのエントランスホール内に常駐しております受付スタッフにお申し付けください。\r\n4.お問い合わせ先　企画部○○課　担当　○○○○まで\r\n　　電話　03-0000-0000";

                kryptonPanel17.Hide();
            }
            else if (treeView2.SelectedNode == miniTreeNode29)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "参加報告書（会議参加報告書）";
                kryptonLabel35.Text = "宛先に対してイベントの参加情報を知らせるための文書です。";

                kryptonTextBox27.Text = "下記のとおり会議に出席しましたので、結果を報告します。";
                kryptonTextBox28.Text = "・日時　○○年〇月○日(月)\r\n・場所　本社〇階　会議室\r\n・出席者　○○部長(B)、○○議長(G)、○○課員(K)、○○名、他3名、合計○○名\r\n・会議の決定事項\r\n   ・○月○日(月)から○月○日(火)まで○○とする\r\n・経過\r\n ・○○を実施する目的として○○の○○を行う必要があることにより、上記のように決定した。";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode30)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "出張報告書";
                kryptonLabel35.Text = "宛先に対して出張の報告を知らせるための文書です。";

                kryptonTextBox27.Text = "下記のとおり出張しましたので、結果を報告します。";
                kryptonTextBox28.Text = "・日時　○○年〇月○日(月)\r\n・場所　○○株式会社　○〇階　会議室 (東京都板橋本町○○丁目)\r\n・内容\r\n      ○○株式会社を行い、下記を行いました。\r\n     ・○○の立会い\r\n・成果\r\n    ・○○を立会いを行い○○の遂行を完了した。\r\n・所感\r\n    ○○の認識が不足していた\r\n・経費\r\n     新幹線代:JT○○　○○線　○○駅～○○駅　○○○○円\r\n　   宿泊代：○○ホテル　(東京都板橋本町○○丁目○○)　○○○○円\r\n\r\n　　承認　　 承認　　承認　　";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode31)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "上申書";
                kryptonLabel35.Text = "上司や上の機関に意見や報告を行う文書です。";

                kryptonTextBox27.Text = "〇〇の件について、下記に記したとおり上申をいたします。\r\n何卒、ご検討のほど宜しくお願い申し上げます。";
                kryptonTextBox28.Text = "（内容を入力）";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode32)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "届出文（○○届）";
                kryptonLabel35.Text = "宛先に対して物品等のお届けをお知らせを行う文書です。";

                kryptonTextBox27.Text = "下記のとおり○○しましたので、お届けいたします。";
                kryptonTextBox28.Text = "1.\r\n2.\r\n3.\r\n";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode33)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "始末書";
                kryptonLabel35.Text = "業務において問題を起こした際に反省の意を示す文書です。";

                kryptonTextBox27.Text = "　私 ○○○○は 、去る〇〇年〇〇月〇〇日、○○○○株式会社との取引において、○○○○○○○【取引停止の原因】を行うという失態を犯し、 上 記 〇〇〇〇株式会社との取引が停止されるという事態を発生させてしまいました。\r\n　今回の件に関しましては〇〇〇〇〇〇【詳細な状況】となったため、このような不始末を起こすこととなりました。\r\n　会社ならびに関係各位に対しましては、多大なる損害ならびにご迷惑をお掛けいたしましたこと、心よりお詫び申し上げます。今後、このような事態を二度と引き起こさないよう、自らを厳しく律し、誠実な態度で日々の業務にあたってまいることを固くお誓い申し上げます。なお、この件につきましては、就業規則に従い、いかなる処分を受けても異議なく存じます。\r\n　つきましては本始末書をもちまして、ここに深くお詫び申し上げます。";
                kryptonTextBox28.Text = "";

                kryptonPanel17.Show();
            }
            else if (treeView2.SelectedNode == miniTreeNode34)
            {
                kryptonButton8.Show();
                kryptonButton12.Show();

                kryptonLabel34.Text = "理由書";
                kryptonLabel35.Text = "業務において問題を起こした際の理由示す文書です。";

                kryptonTextBox27.Text = "この度、○○年〇月○日の○○の件につきまして、理由書を提出させていただきます。";
                kryptonTextBox28.Text = "(理由を入力)";

                kryptonPanel17.Show();
            }
        }

        private void kryptonTextBox29_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupGallery1_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
        }

        private void kryptonContextMenuItem93_Click(object sender, EventArgs e)
        {
            if (kryptonContextMenuItem93.Checked == true)
            {
                kryptonRibbon.StateCommon.RibbonGeneral.RibbonShape = ComponentFactory.Krypton.Toolkit.PaletteRibbonShape.Office2007;
                Properties.Settings.Default.UseOffice2007RibbonShape = true;
                Properties.Settings.Default.Save();
            }
            else
            {
                kryptonRibbon.StateCommon.RibbonGeneral.RibbonShape = ComponentFactory.Krypton.Toolkit.PaletteRibbonShape.Inherit;
                Properties.Settings.Default.UseOffice2007RibbonShape = false;
                Properties.Settings.Default.Save();
            }

        }

        private void kryptonRibbon_ShowQATCustomizeMenu(object sender, ComponentFactory.Krypton.Toolkit.ContextMenuArgs e)
        {
        }

        private void kryptonButton9_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(kryptonTextBox27.Text);
        }

        private void kryptonButton10_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(kryptonTextBox28.Text);
        }

        //テンプレートの内容を適用し編集画面に戻る
        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            kryptonPanel21.Show();

            if (Properties.Settings.Default.RibbonOrMenuBar == "Ribbon")
            {
                kryptonRibbon.Enabled = true;
                this.AllowFormChrome = true;
                kryptonRibbon.MinimizedMode = false;
            }
            else if (Properties.Settings.Default.RibbonOrMenuBar == "MenuBar")
            {
                menuStripPanel.Enabled = true;
                this.AllowFormChrome = false;
            }

            if (kryptonRibbonGroupButton16.Checked == true)
            {
                kryptonPanel21.Height = 36;
                kryptonRibbonGroupButton16.Checked = true;
            }
            else if (kryptonRibbonGroupButton16.Checked == false)
            {
                kryptonPanel21.Height = 0;
                kryptonRibbonGroupButton16.Checked = false;
            }

            if (kryptonLabel34.Text == "注文書")
            {
                kryptonTextBox10.Text = "注文書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();

            }
            if (kryptonLabel34.Text == "承諾書")
            {
                kryptonTextBox10.Text = "承諾書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "依頼文（○○のお願い）")
            {
                kryptonTextBox10.Text = "○○のお願い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "照会文（○○のご照会のお願い）")
            {
                kryptonTextBox10.Text = "○○のご照会のお願い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "回答文（○○に対するお問い合わせの件(回答)）")
            {
                kryptonTextBox10.Text = "○○に対するお問い合わせの件(回答)";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "催促文（○○についての○○のお願い）")
            {
                kryptonTextBox10.Text = "○○についての○○のお願い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "断り文（○○ご辞退の件）")
            {
                kryptonTextBox10.Text = "○○ご辞退の件";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "交渉文（○○のお願い）")
            {
                kryptonTextBox10.Text = "○○のお願い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "抗議文")
            {
                kryptonTextBox10.Text = "抗議文";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "お詫び文（○○についてのお詫び）")
            {
                kryptonTextBox10.Text = "○○についてのお詫び";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "取り消し文（○○の注文の取り消し）")
            {
                kryptonTextBox10.Text = "○○の注文の取り消し";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "あいさつ文（○○新代表取締役社長の就任のお祝いのご挨拶）")
            {
                kryptonTextBox10.Text = "○○新代表取締役社長の就任のお祝いのご挨拶";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "お祝い文（○○様のご結婚のお祝い）")
            {
                kryptonTextBox10.Text = "○○様のご結婚のお祝い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "招待文（○○の新製品発表会のご案内）")
            {
                kryptonTextBox10.Text = "○○の新製品発表会のご案内";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "お礼文（○○のお礼）")
            {
                kryptonTextBox10.Text = "○○のお礼";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "案内文（○○のご案内）")
            {
                kryptonTextBox10.Text = "○○のご案内";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "通知文（○○の通知）")
            {
                kryptonTextBox10.Text = "○○の通知";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "年賀文（あけましておめでとうございます）")
            {
                kryptonTextBox10.Text = "あけましておめでとうございます";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "見舞い文（ご挨拶）")
            {
                kryptonTextBox10.Text = "ご挨拶";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "個人宛見舞い文（暑中お見舞い申し上げます）")
            {
                kryptonTextBox10.Text = "暑中お見舞い申し上げます";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "お悔やみ文（ご遺族の方へ）")
            {
                kryptonTextBox10.Text = "ご遺族の方へ";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "通達文（○○に関する○○のお願い）")
            {
                kryptonTextBox10.Text = "○○に関する○○のお願い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "指示文（○○に関する○○のお願い）")
            {
                kryptonTextBox10.Text = "○○に関する○○のお願い";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "参加報告書（会議参加報告書）")
            {
                kryptonTextBox10.Text = "会議参加報告書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "出張報告書")
            {
                kryptonTextBox10.Text = "出張報告書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "上申書")
            {
                kryptonTextBox10.Text = "上申書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "届出文（○○届）")
            {
                kryptonTextBox10.Text = "○○届";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "始末書")
            {
                kryptonTextBox10.Text = "始末書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }
            else if (kryptonLabel34.Text == "理由書")
            {
                kryptonTextBox10.Text = "理由書";

                Random random = new Random();
                int randomResult = random.Next(1, 9);
                String Title = kryptonTextBox10.Text;
                if (randomResult == 1)
                {
                    kryptonRibbonRecentDoc1.Text = Title;
                    Properties.Settings.Default.RecentDoc1 = Title;
                }
                else if (randomResult == 2)
                {
                    kryptonRibbonRecentDoc2.Text = Title;
                    Properties.Settings.Default.RecentDoc2 = Title;
                }
                else if (randomResult == 3)
                {
                    kryptonRibbonRecentDoc3.Text = Title;
                    Properties.Settings.Default.RecentDoc3 = Title;
                }
                else if (randomResult == 4)
                {
                    kryptonRibbonRecentDoc4.Text = Title;
                    Properties.Settings.Default.RecentDoc4 = Title;
                }
                else if (randomResult == 5)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc5 = Title;
                }
                else if (randomResult == 6)
                {
                    kryptonRibbonRecentDoc6.Text = Title;
                    Properties.Settings.Default.RecentDoc6 = Title;
                }
                else if (randomResult == 7)
                {
                    kryptonRibbonRecentDoc5.Text = Title;
                    Properties.Settings.Default.RecentDoc7 = Title;
                }
                else if (randomResult == 8)
                {
                    kryptonRibbonRecentDoc7.Text = Title;
                    Properties.Settings.Default.RecentDoc8 = Title;
                }
                else if (randomResult == 9)
                {
                    kryptonRibbonRecentDoc8.Text = Title;
                    Properties.Settings.Default.RecentDoc9 = Title;
                }
                Properties.Settings.Default.Save();
            }

            kryptonTextBox12.Text = kryptonTextBox27.Text;
            kryptonTextBox13.Text = kryptonTextBox28.Text;


            //編集画面に戻る
            kryptonPage2.Visible = false;
            kryptonNavigator_Workbench.NavigatorMode = NavigatorMode.BarTabGroup;
            kryptonNavigator_Workbench.SelectedPage = kryptonPage1;

            kryptonRibbon.MinimizedMode = false;
            kryptonRibbon.Enabled = true;

            kryptonLabel7.Enabled = true;
            kryptonCheckButton1.Enabled = true;
            kryptonCheckButton2.Enabled = true;
            kryptonLabel1.Enabled = true;

            if (kryptonTrackBar1.Value == 0)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = false;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }
            else if (kryptonTrackBar1.Value == 10)
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = false;
                kryptonLabel42.Enabled = true;
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                kryptonButton15.Enabled = true;
                kryptonButton14.Enabled = true;
                kryptonLabel42.Enabled = true;
            }

            if (kryptonTextBox10.Text != string.Empty)
            {
                Sheets_TitleButton.Text = kryptonTextBox10.Text;
                this.Text = kryptonTextBox10.Text + " - DocuEase Designer";
            }
            else
            {
                this.Text = "無題 - DocuEase Designer";
            }



            kryptonPage1.AutoScrollPosition = new System.Drawing.Point(0, 0);

            Sheets_Sheet.Anchor = AnchorStyles.Top;

            // 親コントロールのサイズを取得
            int parentWidth = this.ClientSize.Width;
            int parentHeight = this.ClientSize.Height;

            // パネルのサイズを取得
            int panelWidth = Sheets_Sheet.Width;
            int panelHeight = Sheets_Sheet.Height;

            // パネルの位置を中央に設定
            Sheets_Sheet.Location = new System.Drawing.Point(
                (parentWidth - panelWidth) / 2 - 10,
                90
            );

            Sheets_Sheet.Top = 59;
        }

        private void kryptonButton12_Click(object sender, EventArgs e)
        {

            if (kryptonLabel34.Text == "注文書")
            {
                String Title = "注文書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "承諾書")
            {
                String Title = "承諾書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "依頼文（○○のお願い）")
            {
                String Title = "依頼文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "照会文（○○のご照会のお願い）")
            {
                String Title = "照会文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "回答文（○○に対するお問い合わせの件(回答)）")
            {
                String Title = "回答文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "催促文（○○についての○○のお願い）")
            {
                String Title = "催促文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "断り文（○○ご辞退の件）")
            {
                String Title = "断り文文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "交渉文（○○のお願い）")
            {
                String Title = "交渉文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "抗議文")
            {
                String Title = "抗議文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "お詫び文（○○についてのお詫び）")
            {
                String Title = "お詫び文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "取り消し文（○○の注文の取り消し）")
            {
                String Title = "取り消し文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "あいさつ文（○○新代表取締役社長の就任のお祝いのご挨拶）")
            {
                String Title = "あいさつ文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "お祝い文（○○様のご結婚のお祝い）")
            {
                String Title = "お祝い文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "招待文（○○の新製品発表会のご案内）")
            {
                String Title = "招待文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "お礼文（○○のお礼）")
            {
                String Title = "お礼文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));

            }
            else if (kryptonLabel34.Text == "案内文（○○のご案内）")
            {
                String Title = "案内文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "通知文（○○の通知）")
            {
                String Title = "通知文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "年賀文（あけましておめでとうございます）")
            {
                String Title = "年賀文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "見舞い文（ご挨拶）")
            {
                String Title = "見舞い文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "個人宛見舞い文（暑中お見舞い申し上げます）")
            {
                String Title = "個人宛見舞い文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "お悔やみ文（ご遺族の方へ）")
            {
                String Title = "お悔やみ文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "通達文（○○に関する○○のお願い）")
            {
                String Title = "通達文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "指示文（○○に関する○○のお願い）")
            {
                String Title = "指示文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "参加報告書（会議参加報告書）")
            {
                String Title = "参加報告書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "出張報告書")
            {
                String Title = "出張報告書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "上申書")
            {
                String Title = "上申書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "届出文（○○届）")
            {
                String Title = "届出文";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "始末書")
            {
                String Title = "始末書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }
            else if (kryptonLabel34.Text == "理由書")
            {
                String Title = "理由書";
                string searchUrl = "https://www.google.com/search?q=" + Uri.EscapeDataString(Title);
                Process.Start(new ProcessStartInfo(searchUrl));
            }



        }

        private void kryptonRibbonRecentDoc1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc2_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc3_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc4_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc5_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc6_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc7_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc8_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonRecentDoc9_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupGallery1_TrackingImage(object sender, ComponentFactory.Krypton.Toolkit.ImageSelectEventArgs e)
        {

        }

        private void kryptonRibbonGroupGallery1_TrackingImage(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupGallery1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //フォントをリセット
            kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
            kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;
            kryptonRibbonButton_Bold.Checked = false;
            kryptonRibbonButton_Italic.Checked = false;
            kryptonContextMenuItem15.Checked = false;
            kryptonContextMenuItem16.Checked = false;



            //選択されたアイテムに応じてフォントを変更
            if (kryptonRibbonGroupGallery1.SelectedIndex == 0)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;

                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();


            }
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 1)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("Yu Gothic UI", 15, FontStyle.Regular);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.FromArgb(0, 142, 197);

                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 2)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 15, FontStyle.Regular);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 3)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 15, FontStyle.Bold);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                kryptonRibbonButton_Bold.Checked = true;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 4)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Italic);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = true;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 5)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Bold);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                kryptonRibbonButton_Bold.Checked = true;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }
            else if (kryptonRibbonGroupGallery1.SelectedIndex == 6)
            {
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = true;

                Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();
            }
        }

        private void kryptonRibbonGroupButton23_Click(object sender, EventArgs e)
        {
            ImmersiveReaderWindow immersiveReader = new ImmersiveReaderWindow();

            immersiveReader.IsRtfRead = false;
            if (immersiveReader.IsRtfRead == false)
            {
                //シートの内容をイマーシブリーダーウィンドウに送信
                //発行番号
                if (Sheets_NumberLabel.Text != string.Empty)
                {
                    immersiveReader.IssueNumber = Sheets_NumberLabel.Text;
                }
                //日付
                if (Sheets_DateLabel.Text != string.Empty)
                {
                    immersiveReader.Date = Sheets_DateLabel.Text;
                }
                //相手先会社名
                if (Sheets_AddressCompanyLabel.Text != string.Empty)
                {
                    immersiveReader.AdCompany = Sheets_AddressCompanyLabel.Text;
                }
                //相手先氏名
                if (Sheets_AddressTitleAndNameLabel.Text != string.Empty)
                {
                    immersiveReader.AdName = Sheets_AddressTitleAndNameLabel.Text;
                }
                //発信者会社名
                if (Sheets_CallerCompanyLabel.Text != string.Empty)
                {
                    immersiveReader.CaCampany = Sheets_CallerCompanyLabel.Text;
                }
                //発信者所在地
                if (Sheets_CallerLocationLabel.Text != string.Empty)
                {
                    immersiveReader.CaLocation = Sheets_CallerLocationLabel.Text;
                }
                //発信者建物名と階数
                if (Sheets_BuildingNameLabel.Text != string.Empty)
                {
                    immersiveReader.CaBuildingName = Sheets_BuildingNameLabel.Text;
                }
                //発信者氏名
                if (Sheets_CallerTitleAndNameLabel.Text != string.Empty)
                {
                    immersiveReader.CaName = Sheets_CallerTitleAndNameLabel.Text;
                }
                //メールアドレス
                if (Sheets_CallerMallAddressLabel.Text != string.Empty)
                {
                    immersiveReader.CaMailAddress = Sheets_CallerMallAddressLabel.Text;
                }
                //電話番号
                if (Sheets_CallerTelLabel.Text != string.Empty)
                {
                    immersiveReader.CaPhoneNumber1 = Sheets_CallerTelLabel.Text;
                }
                //Fax番号
                if (Sheets_CallerFaxTelLabel.Text != string.Empty)
                {
                    immersiveReader.CaFaxNumber1 = Sheets_CallerFaxTelLabel.Text;
                }
                //表題
                if (Sheets_TitleButton.Text != string.Empty)
                {
                    immersiveReader.title = Sheets_TitleButton.Text;
                }
                //あいさつ文
                if (Sheets_ContentLabel.Text != string.Empty)
                {
                    immersiveReader.Greeting = Sheets_ContentLabel.Text;
                }
                //内容
                if (kryptonTextBox12.Text != string.Empty)
                {
                    immersiveReader.Content = kryptonTextBox12.Text;
                }
                //結語
                if (Sheet_ConclusionLabel.Text != string.Empty)
                {
                    immersiveReader.Conclusion = Sheet_ConclusionLabel.Text;
                }
                //記
                if (Sheet_NoteLabel.Text != string.Empty)
                {
                    immersiveReader.Note = Sheet_NoteLabel.Text;
                }
                //記し書き
                if (kryptonTextBox13.Text != string.Empty)
                {
                    immersiveReader.Notetaking = kryptonTextBox13.Text;
                }
                //以上
                if (Sheets_EndLabel.Text != string.Empty)
                {
                    immersiveReader.Notetaking_End = Sheets_EndLabel.Text;
                }
            }

            //ダイアログの外観設定
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            immersiveReader.ShowDialog();

        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.ForeColor == SystemColors.GrayText)
            {
                toolStripTextBox1.Text = string.Empty;
                toolStripTextBox1.ForeColor = SystemColors.ControlText;
            }
        }


        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("sa");
        }

        private void メニューバーからリボンに切り替えToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //リボンに切り替え
            menuStripPanel.Visible = false;
            kryptonRibbon.Visible = true;

            this.AllowFormChrome = true;

            kryptonContextMenuItem96.Checked = false;

            //次回起動時リボンを使用する
            Properties.Settings.Default.RibbonOrMenuBar = "Ribbon";
            Properties.Settings.Default.Save();

            //リボンにもメニューバーの各値にも適用する

            //1.シート
            //シートの文字色
            kryptonRibbonColorButton_TextColor.SelectedColor = Sheets_SelectForeColor;

            //シートの太字
            if (toolStripButton4.Checked == true)
            {
                kryptonRibbonButton_Bold.Checked = true;
            }
            else if (toolStripButton4.Checked == false)
            {
                kryptonRibbonButton_Bold.Checked = false;
            }

            //シートの斜体
            if (toolStripButton5.Checked == true)
            {
                kryptonRibbonButton_Italic.Checked = true;
            }
            else if (toolStripButton5.Checked == false)
            {
                kryptonRibbonButton_Italic.Checked = false;
            }

            //シートの下線
            if (下線ToolStripMenuItem.Checked == true)
            {
                kryptonContextMenuItem15.Checked = true;
            }
            else if (下線ToolStripMenuItem.Checked == false)
            {
                kryptonContextMenuItem15.Checked = false;
            }

            //シートの打ち消し線
            if (取り消し線ToolStripMenuItem.Checked == true)
            {
                kryptonContextMenuItem16.Checked = true;
            }
            else if (取り消し線ToolStripMenuItem.Checked == false)
            {
                kryptonContextMenuItem16.Checked = false;
            }

            //置換機能の表示状態
            if (置換ToolStripMenuItem.Checked == true)
            {
                kryptonRibbonGroupButton16.Checked = true;
            }
            else if (置換ToolStripMenuItem.Checked == false)
            {
                kryptonRibbonGroupButton16.Checked = false;

            }

            //メモのフォント
            kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;
            kryptonRibbonGroupComboBox_NotepadFontSize.Text = toolStripComboBox4.Text;

            //メモの太字
            if (toolStripButton18.Checked == true)
            {
                kryptonRibbonGroupClusterButton1.Checked = true;
            }
            else if (toolStripButton18.Checked == false)
            {
                kryptonRibbonGroupClusterButton1.Checked = false;
            }

            //メモの斜体
            if (toolStripButton19.Checked == true)
            {
                kryptonRibbonGroupClusterButton2.Checked = true;
            }
            else if (toolStripButton19.Checked == false)
            {
                kryptonRibbonGroupClusterButton2.Checked = false;
            }

            //メモの下線
            if (toolStripMenuItem2.Checked == true)
            {
                kryptonContextMenuItem35.Checked = true;
            }
            else if (toolStripMenuItem2.Checked == false)
            {
                kryptonContextMenuItem35.Checked = false;
            }

            //メモの打ち消し線
            if (toolStripMenuItem3.Checked == true)
            {
                kryptonContextMenuItem36.Checked = true;
            }
            else if (toolStripMenuItem3.Checked == false)
            {
                kryptonContextMenuItem36.Checked = false;
            }

        }

        private void kryptonContextMenuItem96_Click(object sender, EventArgs e)
        {
            menuStripPanel.Visible = true;
            kryptonRibbon.Visible = false;

            this.AllowFormChrome = false;

            kryptonContextMenuItem96.Checked = true;

            Properties.Settings.Default.RibbonOrMenuBar = "MenuBar";
            Properties.Settings.Default.Save();

            //置換機能の表示状態
            if (kryptonRibbonGroupButton16.Checked == true)
            {
                置換ToolStripMenuItem.Checked = true;
            }
            else if (kryptonRibbonGroupButton16.Checked == false)
            {
                置換ToolStripMenuItem.Checked = false;

            }
        }





        private void kryptonScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }


        public void FontReset_Sheets_ToolBar()
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(toolStripComboBox2.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    FontStyle.Regular
                );
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;

            }
        }

        public void FontReset_Notepad_ToolBar()
        {

        }

        private void toolStripComboBox1_DropDownClosed(object sender, EventArgs e)
        {
            Sheets_TitleButton.Font = new System.Drawing.Font(toolStripComboBox1.Text, Sheets_TitleButton.Font.Size, Sheets_TitleButton.Font.Style);
            kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            kryptonRibbonGroupComboBox_Font.Text = toolStripComboBox1.Text;
        }

        private void toolStripComboBox2_DropDownClosed(object sender, EventArgs e)
        {
            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(toolStripComboBox2.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Sheets_TitleButton.Font = new System.Drawing.Font(
                    Sheets_TitleButton.Font.Name,
                    fontSize,
                    Sheets_TitleButton.Font.Style
                );
                kryptonRibbonGroupComboBox_FontSize.Text = fontSize.ToString();
                kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            //太字
            //太字がオンの場合
            if (toolStripButton4.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Bold
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    toolStripButton4.Checked = true;

                }
            }
            //太字がオフの場合
            else if (toolStripButton4.Checked == false)
            {
                //フォントリセット
                FontReset_Sheets_ToolBar();

                float fontSize;
                //斜体が有効な場合
                if (toolStripButton5.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        toolStripButton5.Checked = true;
                    }
                }

                //下線が有効な場合
                if (下線ToolStripMenuItem.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        下線ToolStripMenuItem.Checked = true;
                    }
                }

                //打ち消し線(取り消し線)が有効な場合
                if (取り消し線ToolStripMenuItem.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        取り消し線ToolStripMenuItem.Checked = true;
                    }
                }
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            //斜体
            //斜体がオンの場合
            if (toolStripButton5.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Italic
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    toolStripButton5.Checked = true;
                }
            }
            //斜体がオフの場合
            else if (toolStripButton5.Checked == false)
            {
                //フォントリセット
                FontReset_Sheets_ToolBar();

                float fontSize;
                //太字がオンの場合
                if (toolStripButton4.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        toolStripButton4.Checked = true;

                    }
                }


                //下線が有効な場合
                if (下線ToolStripMenuItem.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        下線ToolStripMenuItem.Checked = true;
                    }
                }

                //打ち消し線(取り消し線)が有効な場合
                if (取り消し線ToolStripMenuItem.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        取り消し線ToolStripMenuItem.Checked = true;
                    }
                }
            }
        }

        private void 下線ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //下線
            //下線がオンの場合
            if (下線ToolStripMenuItem.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Underline
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    下線ToolStripMenuItem.Checked = true;
                }
            }
            else if (下線ToolStripMenuItem.Checked == false)
            {
                //フォントリセット
                FontReset_Sheets_ToolBar();

                float fontSize;
                //太字が有効な場合
                if (toolStripButton4.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        toolStripButton4.Checked = true;

                    }
                }

                //斜体が有効な場合
                if (toolStripButton5.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        toolStripButton5.Checked = true;
                    }
                }

                //取り消し線が有効な場合
                if (取り消し線ToolStripMenuItem.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        取り消し線ToolStripMenuItem.Checked = true;
                    }
                }
            }
        }

        private void 取り消し線ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //取り消し線が有効な場合
            if (取り消し線ToolStripMenuItem.Checked == true)
            {
                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Sheets_TitleButton.Font = new System.Drawing.Font(
                        Sheets_TitleButton.Font.Name,
                        fontSize,
                        Sheets_TitleButton.Font.Style | FontStyle.Strikeout
                    );
                    kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                    取り消し線ToolStripMenuItem.Checked = true;
                }
            }
            else if (取り消し線ToolStripMenuItem.Checked == false)
            {
                //フォントリセット
                FontReset_Sheets_ToolBar();

                float fontSize;
                //太字が有効な場合
                if (toolStripButton4.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Bold
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        toolStripButton4.Checked = true;

                    }
                }

                //斜体が有効な場合
                if (toolStripButton5.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Italic
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        toolStripButton5.Checked = true;
                    }
                }

                //下線が有効な場合
                if (下線ToolStripMenuItem.Checked == true)
                {
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox2.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Sheets_TitleButton.Font = new System.Drawing.Font(
                            Sheets_TitleButton.Font.Name,
                            fontSize,
                            Sheets_TitleButton.Font.Style | FontStyle.Underline
                        );
                        kryptonTextBox10.StateCommon.Content.Font = Sheets_TitleButton.Font;
                        下線ToolStripMenuItem.Checked = true;
                    }
                }
            }
        }

        public Color Sheets_SelectForeColor { get; set; }
        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            using (KryptonColorDialog dialog = new KryptonColorDialog() { Color = Sheets_TitleButton.ForeColor, AnyColor = true })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Sheets_SelectForeColor = dialog.Color;
                    Sheets_TitleButton.ForeColor = dialog.Color;
                    kryptonTextBox10.StateCommon.Content.Color1 = dialog.Color;
                }
            }
        }

        public void FontSizeApply()
        {

        }
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            if (toolStripComboBox2.Text == "8")
            {
                toolStripComboBox2.Text = "9";
            }
            else if (toolStripComboBox2.Text == "9")
            {
                toolStripComboBox2.Text = "10";
            }
            else if (toolStripComboBox2.Text == "9")
            {
                toolStripComboBox2.Text = "10";
            }
            else if (toolStripComboBox2.Text == "10")
            {
                toolStripComboBox2.Text = "10.5";
            }
            else if (toolStripComboBox2.Text == "10.5")
            {
                toolStripComboBox2.Text = "11";
            }
            else if (toolStripComboBox2.Text == "11")
            {
                toolStripComboBox2.Text = "12";
            }
            else if (toolStripComboBox2.Text == "12")
            {
                toolStripComboBox2.Text = "14";
            }
            else if (toolStripComboBox2.Text == "14")
            {
                toolStripComboBox2.Text = "16";
            }
            else if (toolStripComboBox2.Text == "16")
            {
                toolStripComboBox2.Text = "18";
            }
            else if (toolStripComboBox2.Text == "18")
            {
                toolStripComboBox2.Text = "20";
            }
            else if (toolStripComboBox2.Text == "20")
            {
                toolStripComboBox2.Text = "22";
            }
            else if (toolStripComboBox2.Text == "22")
            {
                toolStripComboBox2.Text = "24";
            }
            else if (toolStripComboBox2.Text == "24")
            {
                toolStripComboBox2.Text = "26";
            }
            else if (toolStripComboBox2.Text == "26")
            {
                toolStripComboBox2.Text = "28";
            }
            else if (toolStripComboBox2.Text == "28")
            {
                toolStripComboBox2.Text = "36";
            }
            else if (toolStripComboBox2.Text == "36")
            {
                toolStripComboBox2.Text = "48";
            }
            else if (toolStripComboBox2.Text == "48")
            {
                toolStripComboBox2.Text = "72";
            }
            else if (toolStripComboBox2.Text == "72")
            {
                toolStripComboBox2.Text = "72";
            }
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            if (toolStripComboBox2.Text == "8")
            {
                toolStripComboBox2.Text = "8";
            }
            else if (toolStripComboBox2.Text == "9")
            {
                toolStripComboBox2.Text = "8";
            }
            else if (toolStripComboBox2.Text == "10")
            {
                toolStripComboBox2.Text = "9";
            }
            else if (toolStripComboBox2.Text == "10.5")
            {
                toolStripComboBox2.Text = "10";
            }
            else if (toolStripComboBox2.Text == "11")
            {
                toolStripComboBox2.Text = "10.5";
            }
            else if (toolStripComboBox2.Text == "12")
            {
                toolStripComboBox2.Text = "11";
            }
            else if (toolStripComboBox2.Text == "14")
            {
                toolStripComboBox2.Text = "12";
            }
            else if (toolStripComboBox2.Text == "16")
            {
                toolStripComboBox2.Text = "14";
            }
            else if (toolStripComboBox2.Text == "18")
            {
                toolStripComboBox2.Text = "16";
            }
            else if (toolStripComboBox2.Text == "20")
            {
                toolStripComboBox2.Text = "18";
            }
            else if (toolStripComboBox2.Text == "22")
            {
                toolStripComboBox2.Text = "20";
            }
            else if (toolStripComboBox2.Text == "24")
            {
                toolStripComboBox2.Text = "22";
            }
            else if (toolStripComboBox2.Text == "26")
            {
                toolStripComboBox2.Text = "24";
            }
            else if (toolStripComboBox2.Text == "28")
            {
                toolStripComboBox2.Text = "26";
            }
            else if (toolStripComboBox2.Text == "36")
            {
                toolStripComboBox2.Text = "28";
            }
            else if (toolStripComboBox2.Text == "48")
            {
                toolStripComboBox2.Text = "36";
            }
            else if (toolStripComboBox2.Text == "72")
            {
                toolStripComboBox2.Text = "48";
            }
        }

        private void 閉じるToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void 表題のフォントToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KryptonFontDialog fd = new KryptonFontDialog();
            fd.DisplayExtendedColorsButton = true;
            fd.Font = Sheets_TitleButton.Font;
            fd.ShowColor = true;
            fd.Color = Sheets_TitleButton.ForeColor;
            kryptonRibbonColorButton_TextColor.SelectedColor = Sheets_TitleButton.ForeColor;

            if (fd.ShowDialog() == DialogResult.OK)
            {
                Sheets_TitleButton.Font = fd.Font;
                Sheets_TitleButton.ForeColor = fd.Color;
                Sheets_SelectForeColor = fd.Color;

                kryptonTextBox10.StateCommon.Content.Font = fd.Font;
                kryptonTextBox10.StateCommon.Content.Color1 = fd.Color;

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Bold)
                {
                    toolStripButton4.Checked = true;
                }
                else
                {
                    toolStripButton4.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Underline)
                {
                    下線ToolStripMenuItem.Checked = true;
                }
                else
                {
                    下線ToolStripMenuItem.Checked = false;
                }

                if (kryptonTextBox10.StateCommon.Content.Font.Style == FontStyle.Strikeout)
                {
                    取り消し線ToolStripMenuItem.Checked = true;
                }
                else
                {
                    取り消し線ToolStripMenuItem.Checked = false;
                }
            }
        }


        private void 表題のスタイルToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TitleStyleDialog titleStyleDialog = new TitleStyleDialog();

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                titleStyleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                titleStyleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                titleStyleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                titleStyleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                titleStyleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                titleStyleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            titleStyleDialog.ShowDialog();

            if (titleStyleDialog.DialogResult == DialogResult.OK)
            {
                //フォントをリセット
                kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;
                kryptonRibbonButton_Bold.Checked = false;
                kryptonRibbonButton_Italic.Checked = false;
                kryptonContextMenuItem15.Checked = false;
                kryptonContextMenuItem16.Checked = false;



                //選択されたものに応じてフォントを変更
                if (titleStyleDialog.SetTileStyle == "Defalt")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Regular);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;

                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 0;


                }
                else if (titleStyleDialog.SetTileStyle == "Headline")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("Yu Gothic UI", 15, FontStyle.Regular);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.FromArgb(0, 142, 197);

                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 1;
                }
                else if (titleStyleDialog.SetTileStyle == "Modern")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 15, FontStyle.Regular);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 2;
                }
                else if (titleStyleDialog.SetTileStyle == "ModernBold")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("Segoe UI", 15, FontStyle.Bold);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                    kryptonRibbonButton_Bold.Checked = true;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 3;
                }
                else if (titleStyleDialog.SetTileStyle == "Note")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Italic);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = true;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 4;
                }
                else if (titleStyleDialog.SetTileStyle == "Emphasis")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Bold);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                    kryptonRibbonButton_Bold.Checked = true;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = false;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 5;
                }
                else if (titleStyleDialog.SetTileStyle == "Cancel")
                {
                    kryptonTextBox10.StateCommon.Content.Font = new System.Drawing.Font("游明朝", 12, FontStyle.Strikeout);
                    kryptonTextBox10.StateCommon.Content.Color1 = Color.Black;

                    kryptonRibbonButton_Bold.Checked = false;
                    kryptonRibbonButton_Italic.Checked = false;
                    kryptonContextMenuItem15.Checked = false;
                    kryptonContextMenuItem16.Checked = true;

                    Sheets_TitleButton.Font = kryptonTextBox10.StateCommon.Content.Font;
                    Sheets_TitleButton.ForeColor = kryptonTextBox10.StateCommon.Content.Color1;


                    kryptonRibbonGroupComboBox_Font.Text = Sheets_TitleButton.Font.Name;
                    kryptonRibbonGroupComboBox_FontSize.Text = Sheets_TitleButton.Font.Size.ToString();

                    kryptonRibbonGroupGallery1.SelectedIndex = 6;
                }
            }
        }

        private void シートの外枠の余白ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //シートの余白設定ダイアログを表示
            SheetSpaceSettingDialog sheetSpaceSettingDialog = new SheetSpaceSettingDialog();

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                sheetSpaceSettingDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                sheetSpaceSettingDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                sheetSpaceSettingDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                sheetSpaceSettingDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                sheetSpaceSettingDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                sheetSpaceSettingDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            //現在の余白設定を送信
            sheetSpaceSettingDialog.TopMargin = Sheets_TopPanel.Height;
            sheetSpaceSettingDialog.ButtomMargin = Sheets_ButtomPanel.Height;
            sheetSpaceSettingDialog.LeftMargin = Sheets_LeftPanel.Width;
            sheetSpaceSettingDialog.RightMargin = Sheets_RightPanel.Width;

            sheetSpaceSettingDialog.ShowDialog();

            if (sheetSpaceSettingDialog.DialogResult == DialogResult.OK)
            {
                Sheets_TopPanel.Height = sheetSpaceSettingDialog.TopMargin;
                Sheets_ButtomPanel.Height = sheetSpaceSettingDialog.TopMargin;
                Sheets_LeftPanel.Width = sheetSpaceSettingDialog.RightMargin;
                Sheets_RightPanel.Width = sheetSpaceSettingDialog.LeftMargin;
            }
        }

        private void 置換ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (置換ToolStripMenuItem.Checked == true)
            {

                Transition
                    .With(kryptonPanel21, nameof(Height), 36)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                置換ToolStripMenuItem.Checked = true;
                kryptonRibbonGroupButton16.Checked = true;

                kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.SlantOutsizeFar;
            }
            else if (置換ToolStripMenuItem.Checked == false)
            {
                kryptonNavigator_Workbench.Bar.TabBorderStyle = ComponentFactory.Krypton.Toolkit.TabBorderStyle.OneNote;
                Transition
                    .With(kryptonPanel21, nameof(Height), 0)
                    .CriticalDamp(TimeSpan.FromSeconds(0.4));
                置換ToolStripMenuItem.Checked = false;
                kryptonRibbonGroupButton16.Checked = false;
            }

        }


        private static Form1 _form1Instance;

        //Form1オブジェクトを取得、設定するためのプロパティ
        public static Form1 Form1Instance
        {
            get
            {
                return _form1Instance;
            }
            set
            {
                _form1Instance = value;
            }
        }

        private void シートのサイズToolStripMenuItem_Click(object sender, EventArgs e)
        {

            シートのサイズToolStripMenuItem.Enabled = false;
            kryptonRibbonGroup6.DialogBoxLauncher = false;

            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                sheetsScaleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                sheetsScaleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                sheetsScaleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                sheetsScaleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                sheetsScaleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                sheetsScaleDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }

            sheetsScaleDialog.Show();
        }

        private void 全画面モードToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (全画面モードToolStripMenuItem.Checked == true)
            {
                this.WindowState = FormWindowState.Normal;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
                this.AllowFormChrome = false;

                全画面モードToolStripMenuItem.Checked = true;
                kryptonRibbonGroupButton2.Checked = true;
            }
            else if (全画面モードToolStripMenuItem.Checked == false)
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.WindowState = FormWindowState.Normal;
                this.AllowFormChrome = false;
                全画面モードToolStripMenuItem.Checked = false;
                kryptonRibbonGroupButton2.Checked = false;
            }
        }

        private void 文書の共有ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShareDialog shareDialog = new ShareDialog();

            if (Sheets_TitleButton.Text != string.Empty)
            {
                shareDialog.ShareTitle = Sheets_TitleButton.Text;
            }
            else
            {
                shareDialog.ShareTitle = "無題の文書";
            }


            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                shareDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                shareDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                shareDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                shareDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                shareDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                shareDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            shareDialog.ShowDialog();
            if (shareDialog.ShareContent == "MicrosoftOutlook")
            {
                MessageBox.Show("Outlook");
            }
            else if (shareDialog.ShareContent == "MicrosoftTeams")
            {
                MessageBox.Show("Teams");
            }
            else if (shareDialog.ShareContent == "Slack")
            {
                MessageBox.Show("Slack");
            }
            shareDialog.Dispose();
        }

        private void ツールToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ソフトウェアの外観とテーマToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeThemeDialog changeThemeDialog = new ChangeThemeDialog();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                changeThemeDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                changeThemeDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                changeThemeDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                changeThemeDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                changeThemeDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                changeThemeDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            changeThemeDialog.ShowDialog();
            if (changeThemeDialog.DialogResult == DialogResult.OK)
            {
                //テーマの変更
                //テーマ取得処理
                //Office2007青の場合
                if (changeThemeDialog.SelectedTheme == "Office2007Blue")
                {
                    Properties.Settings.Default.Theme = "Office2007Blue";
                }
                //Office2007銀色の場合
                else if (changeThemeDialog.SelectedTheme == "Office2007Silver")
                {
                    Properties.Settings.Default.Theme = "Office2007Silver";
                }
                //Office2007黒の場合
                else if (changeThemeDialog.SelectedTheme == "Office2007Black")
                {
                    Properties.Settings.Default.Theme = "Office2007Black";
                }
                //Office2010青の場合
                else if (changeThemeDialog.SelectedTheme == "Office2010Blue")
                {
                    Properties.Settings.Default.Theme = "Office2010Blue";
                }
                //Office2010銀色の場合
                else if (changeThemeDialog.SelectedTheme == "Office2010Silver")
                {
                    Properties.Settings.Default.Theme = "Office2010Silver";
                }
                //Office2010黒の場合
                else if (changeThemeDialog.SelectedTheme == "Office2010Black")
                {
                    Properties.Settings.Default.Theme = "Office2010Black";
                }

                //リボンシェイプ設定を保存
                if (changeThemeDialog.UseOffice2007RibbonMenuAndQAT == true)
                {
                    Properties.Settings.Default.UseOffice2007RibbonShape = true;
                }
                else if (changeThemeDialog.UseOffice2007RibbonMenuAndQAT == false)
                {
                    Properties.Settings.Default.UseOffice2007RibbonShape = false;
                }

                //リボンかメニューバー設定を保存
                if (changeThemeDialog.RibbonOrMenuBar == "Ribbon")
                {
                    Properties.Settings.Default.RibbonOrMenuBar = "Ribbon";
                }
                else if (changeThemeDialog.RibbonOrMenuBar == "MenuBar")
                {
                    Properties.Settings.Default.RibbonOrMenuBar = "MenuBar";
                }

                //設定を保存
                Properties.Settings.Default.Save();

                //設定を適用
                SetTheme();
                SetRibbonOrMenuBar();

                //ガベージコレクション
                GC.Collect();
            }
        }

        private void ファイルToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void kryptonNumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (kryptonTextBox11.Text != string.Empty)
            {
                label2.Text = "発第";
                Sheets_NumberLabel.Text = kryptonTextBox11.Text + "発第" + kryptonNumericUpDown1.Value + "号";
            }
            else
            {
                label2.Text = "　第";
                Sheets_NumberLabel.Text = "第" + kryptonNumericUpDown1.Value + "号";
            }
        }

        //機能検索
        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            //プレースホルダーを削除
            string text = toolStripTextBox1.Text;
            if (toolStripTextBox1.ForeColor == SystemColors.GrayText)
            {
                toolStripTextBox1.Text = string.Empty;
                toolStripTextBox1.ForeColor = SystemColors.ControlText;
                toolStripTextBox1.Text = text;
            }
            else
            {
                //辞書から検索
                //印刷
                if (toolStripTextBox1.Text == "印刷"
                    | toolStripTextBox1.Text == "print"
                    | toolStripTextBox1.Text == "Print")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "印刷";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "印刷プレビューを表示";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //印刷プレビュー
                else if (toolStripTextBox1.Text == "印刷プレビュー"
                    | toolStripTextBox1.Text == "printpreview"
                    | toolStripTextBox1.Text == "PrintPreview"
                    | toolStripTextBox1.Text == "print preview"
                    | toolStripTextBox1.Text == "Print Preview")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "印刷プレビューを表示";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //プレビュー(曖昧検索)
                else if (toolStripTextBox1.Text == "プレビュー"
                    | toolStripTextBox1.Text == "preview"
                    | toolStripTextBox1.Text == "Preview")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "印刷プレビューを表示";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "閲覧用(表示モード)";
                    SearchSuggestions3.Enabled = true;
                    SearchSuggestions3.Text = "全画面モード";
                }
                //文書作成ウィザード
                else if (toolStripTextBox1.Text == "文書作成ウィザード"
                    | toolStripTextBox1.Text == "ウィザード"
                    | toolStripTextBox1.Text == "DocumentCreationWizard"
                    | toolStripTextBox1.Text == "Document Creation Wizard"
                    | toolStripTextBox1.Text == "Wizard"
                    | toolStripTextBox1.Text == "wizard")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "文書作成ウィザードを起動";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //文書作成,文書
                else if (toolStripTextBox1.Text == "文書作成"
                    | toolStripTextBox1.Text == "DocumentCreation"
                    | toolStripTextBox1.Text == "Document Creation"
                    | toolStripTextBox1.Text == "文書"
                    | toolStripTextBox1.Text == "Document"
                    | toolStripTextBox1.Text == "Documents"
                    | toolStripTextBox1.Text == "Doc"
                    | toolStripTextBox1.Text == "Docs")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "文書作成ウィザードを起動";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "文書作成ソフトウェアで編集する";
                    SearchSuggestions3.Enabled = true;
                    SearchSuggestions3.Text = "Docxファイルとして保存";
                }
                //Docxファイルとして保存
                else if (toolStripTextBox1.Text == "Docxファイルとして保存"
                    | toolStripTextBox1.Text == "Docx"
                    | toolStripTextBox1.Text == "Docx Save"
                    | toolStripTextBox1.Text == "Save as a Docx file"
                    | toolStripTextBox1.Text == "saveasaDocxfile"
                    | toolStripTextBox1.Text == "save as a docx file"
                    | toolStripTextBox1.Text == "Save as Docx file"
                    | toolStripTextBox1.Text == "saveasaDocxfile")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "Docxファイルとして保存";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //保存(曖昧検索)
                else if (toolStripTextBox1.Text == "保存"
                    | toolStripTextBox1.Text == "Save"
                    | toolStripTextBox1.Text == "Save as"
                    | toolStripTextBox1.Text == "save as"
                    | toolStripTextBox1.Text == "saveas"
                    | toolStripTextBox1.Text == "save")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "Docxファイルとして保存";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "保存(メモ帳)";
                    SearchSuggestions3.Enabled = true;
                    SearchSuggestions3.Text = "名前を付けてコピーとして保存(メモ帳)";
                }
                //シート
                else if (toolStripTextBox1.Text == "シート"
                    | toolStripTextBox1.Text == "Sheet"
                    | toolStripTextBox1.Text == "sheet")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "シート(タブ)";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "シートの余白(スタイルリボンタブ)";
                    SearchSuggestions3.Enabled = true;
                    SearchSuggestions3.Text = "シートの拡大率ダイアログ";
                }
                //連絡先,連絡帳
                else if (toolStripTextBox1.Text == "連絡先"
                    | toolStripTextBox1.Text == "連絡帳"
                    | toolStripTextBox1.Text == "コンタクト"
                    | toolStripTextBox1.Text == "Contact"
                    | toolStripTextBox1.Text == "Contacts"
                    | toolStripTextBox1.Text == "contact"
                    | toolStripTextBox1.Text == "contacts")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "連絡帳(タブ)";
                    if (kryptonPanel20.Visible == false)
                    {
                        SearchSuggestions2.Enabled = true;
                        SearchSuggestions2.Text = "連絡先を追加";
                        SearchSuggestions3.Enabled = true;
                        SearchSuggestions3.Text = "連絡先を削除";
                    }
                    else
                    {
                        SearchSuggestions2.Enabled = false;
                        SearchSuggestions2.Text = "候補2";
                        SearchSuggestions3.Enabled = false;
                        SearchSuggestions3.Text = "候補3";
                    }

                }
                //メモ帳
                else if (toolStripTextBox1.Text == "メモ帳"
                    | toolStripTextBox1.Text == "メモ"
                    | toolStripTextBox1.Text == "Notepad"
                    | toolStripTextBox1.Text == "notepad"
                    | toolStripTextBox1.Text == "memo"
                    | toolStripTextBox1.Text == "Memo")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "メモ帳(タブ)";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "保存(メモ帳)";
                    SearchSuggestions3.Enabled = true;
                    SearchSuggestions3.Text = "名前を付けてコピーを保存(メモ帳)";
                }
                //リセット,余白をすべてリセット,編集内容をすべてリセット
                else if (toolStripTextBox1.Text == "リセット"
                    | toolStripTextBox1.Text == "Reset"
                    | toolStripTextBox1.Text == "reset"
                    | toolStripTextBox1.Text == "余白をすべてリセット"
                    | toolStripTextBox1.Text == "編集内容をすべてリセット"
                    )
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "余白をすべてリセット";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "編集内容をすべてリセット";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //書式
                else if (toolStripTextBox1.Text == "書式"
                    | toolStripTextBox1.Text == "Paragraph"
                    | toolStripTextBox1.Text == "paragraph"
                    | toolStripTextBox1.Text == "フォント"
                    | toolStripTextBox1.Text == "Font"
                    | toolStripTextBox1.Text == "font")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "表題のフォント(ホームリボンタブ)";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "フォント(メモ帳リボンタブ)";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //シートの余白
                else if (toolStripTextBox1.Text == "シートの余白"
                     | toolStripTextBox1.Text == "余白"
                     | toolStripTextBox1.Text == "間隔"
                     | toolStripTextBox1.Text == "Space"
                     | toolStripTextBox1.Text == "Sheet Space"
                     | toolStripTextBox1.Text == "sheet space"
                     | toolStripTextBox1.Text == "SheetSpace"
                     | toolStripTextBox1.Text == "sheetspace"
                     | toolStripTextBox1.Text == "space"
                     | toolStripTextBox1.Text == "Margin"
                     | toolStripTextBox1.Text == "Sheet Margin"
                     | toolStripTextBox1.Text == "sheet margin"
                     | toolStripTextBox1.Text == "SheetMargin"
                     | toolStripTextBox1.Text == "sheetmargin"
                     | toolStripTextBox1.Text == "margin")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "シートの余白(スタイルリボンタブ)";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "余白をすべてリセット";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //シートの拡大率
                else if (toolStripTextBox1.Text == "拡大率"
                    | toolStripTextBox1.Text == "スケール"
                    | toolStripTextBox1.Text == "シートの拡大率")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "シートの拡大率(タブ)";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "シートの拡大率ダイアログ";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //スタイル(曖昧検索)
                else if (toolStripTextBox1.Text == "表題のスタイル"
                    | toolStripTextBox1.Text == "スタイル"
                    | toolStripTextBox1.Text == "Style"
                    | toolStripTextBox1.Text == "style")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "スタイル(ホームリボンタブ)";
                    SearchSuggestions2.Enabled = true;
                    SearchSuggestions2.Text = "表題のスタイル";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //全画面モード
                else if (toolStripTextBox1.Text == "全画面モード"
                    | toolStripTextBox1.Text == "フルスクリーンモード"
                    | toolStripTextBox1.Text == "全画面"
                    | toolStripTextBox1.Text == "フルスクリーン"
                    | toolStripTextBox1.Text == "Full Screen Mode"
                    | toolStripTextBox1.Text == "FullScreenMode"
                    | toolStripTextBox1.Text == "full screen mode"
                    | toolStripTextBox1.Text == "fullscreensode"
                    | toolStripTextBox1.Text == "Full screen mode")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "全画面モードの切り替え";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //テンプレートを選択
                else if (toolStripTextBox1.Text == "テンプレートを選択"
                    | toolStripTextBox1.Text == "テンプレートの選択"
                    | toolStripTextBox1.Text == "Select Template"
                    | toolStripTextBox1.Text == "SelectTemplate"
                    | toolStripTextBox1.Text == "select template"
                    | toolStripTextBox1.Text == "selecttemplate"
                    | toolStripTextBox1.Text == "Select template"
                    | toolStripTextBox1.Text == "Select Templates"
                    | toolStripTextBox1.Text == "SelectTemplates"
                    | toolStripTextBox1.Text == "select templates"
                    | toolStripTextBox1.Text == "selecttemplates"
                    | toolStripTextBox1.Text == "Select templates"
                    | toolStripTextBox1.Text == "テンプレート"
                    | toolStripTextBox1.Text == "Template"
                    | toolStripTextBox1.Text == "template"
                    | toolStripTextBox1.Text == "Templates"
                    | toolStripTextBox1.Text == "templates")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "テンプレート選択画面";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //ウィンドウ
                else if (toolStripTextBox1.Text == "ウィンドウ"
                    | toolStripTextBox1.Text == "新しいウィンドウ"
                    | toolStripTextBox1.Text == "Window"
                    | toolStripTextBox1.Text == "window"
                    | toolStripTextBox1.Text == "New Window"
                    | toolStripTextBox1.Text == "new window"
                    | toolStripTextBox1.Text == "newwindow"
                    | toolStripTextBox1.Text == "New window")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "新しいウィンドウを表示";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //イマーシブリーダー
                else if (toolStripTextBox1.Text == "イマーシブリーダー"
                    | toolStripTextBox1.Text == "リーダー"
                    | toolStripTextBox1.Text == "Reader"
                    | toolStripTextBox1.Text == "Reader"
                    | toolStripTextBox1.Text == "Immersive Reader"
                    | toolStripTextBox1.Text == "immersive reader"
                    | toolStripTextBox1.Text == "immersivereader"
                    | toolStripTextBox1.Text == "Immersive reader")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "イマーシブリーダー";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //設定
                else if (toolStripTextBox1.Text == "設定"
                    | toolStripTextBox1.Text == "オプション"
                    | toolStripTextBox1.Text == "Setting"
                    | toolStripTextBox1.Text == "setting"
                    | toolStripTextBox1.Text == "Option"
                    | toolStripTextBox1.Text == "option")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "設定";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //テーマ設定
                else if (toolStripTextBox1.Text == "テーマ"
                    | toolStripTextBox1.Text == "テーマ設定"
                    | toolStripTextBox1.Text == "Theme"
                    | toolStripTextBox1.Text == "theme"
                    | toolStripTextBox1.Text == "Theme Setting"
                    | toolStripTextBox1.Text == "theme setting"
                    | toolStripTextBox1.Text == "themesetting"
                    | toolStripTextBox1.Text == "Theme setting"
                    | toolStripTextBox1.Text == "ThemeSetting")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "テーマ設定ダイアログ";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
                //置換
                else if (toolStripTextBox1.Text == "置換"
                    | toolStripTextBox1.Text == "Replacement"
                    | toolStripTextBox1.Text == "replacement"
                    | toolStripTextBox1.Text == "Replace"
                    | toolStripTextBox1.Text == "replace")
                {
                    SearchSuggestions1.Enabled = true;
                    SearchSuggestions1.Text = "文字の置換(シート)";
                    SearchSuggestions2.Enabled = false;
                    SearchSuggestions2.Text = "候補2";
                    SearchSuggestions3.Enabled = false;
                    SearchSuggestions3.Text = "候補3";
                }
            }

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void toolStripMenuItem1_MouseEnter(object sender, EventArgs e)
        {
            toolStripMenuItem1.BackColor = Color.White;
        }

        private void kryptonRibbonGroupClusterButton8_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void kryptonRibbonGroupClusterButton9_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void kryptonRibbonGroupClusterButton10_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void kryptonRibbonGroupClusterButton11_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionIndent = Notepads_kryptonRichTextBox_Notepad.SelectionIndent - 10;
        }

        private void kryptonRibbonGroupClusterButton12_Click(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.SelectionIndent = Notepads_kryptonRichTextBox_Notepad.SelectionIndent + 10;

        }

        private void Notepads_kryptonRichTextBox_Notepad_KeyDown(object sender, KeyEventArgs e)
        {


        }

        private void kryptonRibbonGroupClusterButton13_Click(object sender, EventArgs e)
        {
            //箇条書きをする
            Notepads_kryptonRichTextBox_Notepad.SelectionBullet = !Notepads_kryptonRichTextBox_Notepad.SelectionBullet;
            //箇条書き確認
            if (Notepads_kryptonRichTextBox_Notepad.SelectionBullet == true)
            {
                kryptonRibbonGroupClusterButton13.Checked = true;
                toolStripButton39.Checked = true;
            }
            else if (Notepads_kryptonRichTextBox_Notepad.SelectionBullet == false)
            {
                kryptonRibbonGroupClusterButton13.Checked = false;
                toolStripButton39.Checked = false;
            }
        }

        private void kryptonRibbonGroupButton25_Click(object sender, EventArgs e)
        {
            FileDeleteWarningDialog fileDeleteWarningDialog = new FileDeleteWarningDialog();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                fileDeleteWarningDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                fileDeleteWarningDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                fileDeleteWarningDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                fileDeleteWarningDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                fileDeleteWarningDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                fileDeleteWarningDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            fileDeleteWarningDialog.ShowDialog();
            if (fileDeleteWarningDialog.DialogResult == DialogResult.Yes)
            {
                try
                {

                    String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase";


                    if (Directory.Exists(str))
                    {
                        //ファイルを削除しトーストー通知で削除したことを表示する
                        File.Delete(str + @"\SaveFile.rtf");
                        Directory.Delete(str);


                        new ToastContentBuilder()
                            .AddText("自動保存ファイルは正常に削除されました")
                            .AddText("自動保存ファイルは正しく削除されDocuEase\nを終了しました。")
                            .Show();

                        SaveFileDeleted = true;
                        System.Windows.Forms.Application.Exit();
                    }
                }
                catch { }
            }
        }

        private void kryptonRibbonGroupButton24_Click(object sender, EventArgs e)
        {
            AddWebLinkDialog addWebLinkDialog = new AddWebLinkDialog();
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                addWebLinkDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                addWebLinkDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                addWebLinkDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                addWebLinkDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                addWebLinkDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                addWebLinkDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            addWebLinkDialog.ShowDialog();
            if (addWebLinkDialog.DialogResult == DialogResult.OK)
            {
                Notepads_kryptonRichTextBox_Notepad.Text += addWebLinkDialog.WebLink;
            }
        }

        private void kryptonRibbonGroupButton31_Click(object sender, EventArgs e)
        {
            if (Notepads_kryptonRichTextBox_Notepad.SelectionLength > 0)
            {
                string selected = Notepads_kryptonRichTextBox_Notepad.SelectedText;
                string modified = selected.Remove(0, Notepads_kryptonRichTextBox_Notepad.SelectionLength);

                // 修正後の文字列を代入
                Notepads_kryptonRichTextBox_Notepad.SelectedText = modified;
            }

        }

        private void kryptonRibbonGroupButton25_Click_1(object sender, EventArgs e)
        {
            編集用ToolStripMenuItem1.Checked = true;
            閲覧用ToolStripMenuItem1.Checked = false;

            kryptonRibbonGroupButton25.Checked = true;
            kryptonRibbonGroupButton30.Checked = false;

            Notepads_kryptonRichTextBox_Notepad.ReadOnly = false;

            kryptonRibbonGroupButton_NotepadPaste.Enabled = true;
            kryptonRibbonGroupButton_NotepadCopy.Enabled = true;
            kryptonRibbonGroupButton_NotepadCut.Enabled = true;
            kryptonRibbonGroupButton31.Enabled = true;
            kryptonRibbonGroupComboBox_NotepadFont.Enabled = true;
            kryptonRibbonGroupClusterButton1.Enabled = true;
            kryptonRibbonGroupClusterButton2.Enabled = true;
            kryptonRibbonGroupClusterButton3.Enabled = true;
            kryptonRibbonGroupColorButton2.Enabled = true;
            kryptonRibbonGroupColorButton3.Enabled = true;
            kryptonRibbonGroupComboBox_NotepadFontSize.Enabled = true;
            kryptonRibbonGroupClusterButton6.Enabled = true;
            kryptonRibbonGroupClusterButton7.Enabled = true;
            kryptonRibbonGroupClusterButton8.Enabled = true;
            kryptonRibbonGroupClusterButton9.Enabled = true;
            kryptonRibbonGroupClusterButton10.Enabled = true;
            kryptonRibbonGroupClusterButton11.Enabled = true;
            kryptonRibbonGroupClusterButton12.Enabled = true;
            kryptonRibbonGroupClusterButton13.Enabled = true;
            kryptonRibbonGroupGallery2.Enabled = true;
            kryptonRibbonGroupButton28.Enabled = true;
            kryptonRibbonGroupButton29.Enabled = true;
            kryptonRibbonGroupButton5.Enabled = true;
            kryptonRibbonGroupButton8.Enabled = true;
            kryptonContextMenu8.Enabled = true;

            toolStripButton15.Enabled = true;
            toolStripButton16.Enabled = true;
            toolStripButton17.Enabled = true;
            toolStripComboBox3.Enabled = true;
            toolStripComboBox4.Enabled = true;
            toolStripButton18.Enabled = true;
            toolStripButton19.Enabled = true;
            toolStripDropDownButton2.Enabled = true;
            toolStripButton20.Enabled = true;
            toolStripButton21.Enabled = true;
            toolStripButton22.Enabled = true;
            toolStripButton22.Enabled = true;
            toolStripButton34.Enabled = true;
            toolStripButton35.Enabled = true;
            toolStripButton36.Enabled = true;
            toolStripButton37.Enabled = true;
            toolStripButton38.Enabled = true;
            toolStripButton39.Enabled = true;
            toolStripButton27.Enabled = true;
            toolStripButton31.Enabled = true;
            toolStripDropDownButton3.Enabled = true;
            toolStripButton23.Enabled = true;

            kryptonRibbonQATButton4.Enabled = true;
            kryptonRibbonQATButton5.Enabled = true;
            toolStripButton28.Enabled = true;
            toolStripButton29.Enabled = true;
        }

        private void kryptonRibbonGroupButton30_Click(object sender, EventArgs e)
        {
            編集用ToolStripMenuItem1.Checked = false;
            閲覧用ToolStripMenuItem1.Checked = true;

            kryptonRibbonGroupButton25.Checked = false;
            kryptonRibbonGroupButton30.Checked = true;

            Notepads_kryptonRichTextBox_Notepad.ReadOnly = true;

            kryptonRibbonGroupButton_NotepadPaste.Enabled = false;
            kryptonRibbonGroupButton_NotepadCopy.Enabled = false;
            kryptonRibbonGroupButton_NotepadCut.Enabled = false;
            kryptonRibbonGroupButton31.Enabled = false;
            kryptonRibbonGroupComboBox_NotepadFont.Enabled = false;
            kryptonRibbonGroupClusterButton1.Enabled = false;
            kryptonRibbonGroupClusterButton2.Enabled = false;
            kryptonRibbonGroupClusterButton3.Enabled = false;
            kryptonRibbonGroupColorButton2.Enabled = false;
            kryptonRibbonGroupColorButton3.Enabled = false;
            kryptonRibbonGroupComboBox_NotepadFontSize.Enabled = false;
            kryptonRibbonGroupClusterButton6.Enabled = false;
            kryptonRibbonGroupClusterButton7.Enabled = false;
            kryptonRibbonGroupClusterButton8.Enabled = false;
            kryptonRibbonGroupClusterButton9.Enabled = false;
            kryptonRibbonGroupClusterButton10.Enabled = false;
            kryptonRibbonGroupClusterButton11.Enabled = false;
            kryptonRibbonGroupClusterButton12.Enabled = false;
            kryptonRibbonGroupClusterButton13.Enabled = false;
            kryptonRibbonGroupGallery2.Enabled = false;
            kryptonRibbonGroupButton28.Enabled = false;
            kryptonRibbonGroupButton29.Enabled = false;
            kryptonRibbonGroupButton5.Enabled = false;
            kryptonRibbonGroupButton8.Enabled = false;
            kryptonContextMenu8.Enabled = false;

            toolStripButton15.Enabled = false;
            toolStripButton16.Enabled = false;
            toolStripButton17.Enabled = false;
            toolStripComboBox3.Enabled = false;
            toolStripComboBox4.Enabled = false;
            toolStripButton18.Enabled = false;
            toolStripButton19.Enabled = false;
            toolStripDropDownButton2.Enabled = false;
            toolStripButton20.Enabled = false;
            toolStripButton21.Enabled = false;
            toolStripButton22.Enabled = false;
            toolStripButton22.Enabled = false;
            toolStripButton34.Enabled = false;
            toolStripButton35.Enabled = false;
            toolStripButton36.Enabled = false;
            toolStripButton37.Enabled = false;
            toolStripButton38.Enabled = false;
            toolStripButton39.Enabled = false;
            toolStripButton27.Enabled = false;
            toolStripButton31.Enabled = false;
            toolStripDropDownButton3.Enabled = false;
            toolStripButton23.Enabled = false;


            kryptonRibbonQATButton4.Enabled = false;
            kryptonRibbonQATButton5.Enabled = false;
            toolStripButton28.Enabled = false;
            toolStripButton29.Enabled = false;
        }

        private void kryptonRibbonGroupButton33_Click(object sender, EventArgs e)
        {


            try
            {
                String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase";

                if (Directory.Exists(str))
                {
                    using (PrintDialog printDialog = new PrintDialog() { UseEXDialog = true })
                    {
                        if (printDialog.ShowDialog() == DialogResult.OK)
                        {
                            GC.Collect();
                            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                            //バックグラウンド上で起動
                            word.Visible = false;
                            //読み取り専用でrtfファイルを開く
                            word.Documents.Open(str + @"\SaveFile.rtf", ReadOnly: true);
                            //使用するプリンターを設定し印刷
                            word.ActivePrinter = printDialog.PrinterSettings.PrinterName;
                            word.PrintOut();
                            //保存を確認せずwordを閉じる
                            word.Quit();
                            GC.Collect();
                            //一応明示的に破棄
                            printDialog.Dispose();
                        }
                    }

                }
                else
                {
                    PrintErrorDialog printErrorDialog = new PrintErrorDialog();
                    //Office2007青色
                    if (this.BackColor == Color.FromArgb(191, 219, 255))
                    {
                        printErrorDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
                    }
                    //Office2007銀色
                    else if (this.BackColor == Color.FromArgb(208, 212, 221))
                    {
                        printErrorDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
                    }
                    //Office2007ブラック
                    else if (this.BackColor == Color.FromArgb(83, 83, 83))
                    {
                        printErrorDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
                    }
                    //Office2010青色
                    else if (this.BackColor == Color.FromArgb(187, 206, 230))
                    {
                        printErrorDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
                    }
                    //Office2010銀色
                    else if (this.BackColor == Color.FromArgb(227, 230, 232))
                    {
                        printErrorDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
                    }
                    //Office2010黒色
                    else if (this.BackColor == Color.FromArgb(113, 113, 113))
                    {
                        printErrorDialog.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
                    }

                    printErrorDialog.ShowDialog();
                }
            }
            catch { }

        }

        private void kryptonRibbonGroupButton32_Click(object sender, EventArgs e)
        {
            ImmersiveReaderWindow immersiveReader = new ImmersiveReaderWindow();

            immersiveReader.IsRtfRead = true;
            //Rtfの保存処理
            if (immersiveReader.IsRtfRead == true)
            {
                //ファイル保存
                AutoSave();
            }

            //ダイアログの外観設定
            //Office2007青色
            if (this.BackColor == Color.FromArgb(191, 219, 255))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            //Office2007銀色
            else if (this.BackColor == Color.FromArgb(208, 212, 221))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            //Office2007ブラック
            else if (this.BackColor == Color.FromArgb(83, 83, 83))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;
            }
            //Office2010青色
            else if (this.BackColor == Color.FromArgb(187, 206, 230))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            //Office2010銀色
            else if (this.BackColor == Color.FromArgb(227, 230, 232))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            //Office2010黒色
            else if (this.BackColor == Color.FromArgb(113, 113, 113))
            {
                immersiveReader.PaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
            immersiveReader.ShowDialog();
        }

        private void toolStripComboBox3_DropDownClosed(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
            // 現在のフォント名を変更する
            Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                toolStripComboBox3.Text,
                Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style
            );
            kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;
        }

        private void toolStripComboBox3_DropDownClosed(object sender, KeyEventArgs e)
        {

        }

        private void toolStripComboBox4_DropDownClosed(object sender, EventArgs e)
        {
            Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

            float fontSize;
            // 入力値がfloatに変換できるかチェック
            if (float.TryParse(toolStripComboBox4.Text, out fontSize))
            {
                // 現在のフォント名とスタイルを維持し、サイズのみ変更
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                    fontSize,
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style
                );
            }

            kryptonRibbonGroupComboBox_NotepadFontSize.Text = toolStripComboBox4.Text;
        }

        private void toolStripComboBox4_DropDownClosed(object sender, KeyEventArgs e)
        {

        }

        //太字
        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            if (toolStripButton18.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                    );
                }

                //完了後他のフォントスタイルを確認
                toolStripButton18.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    toolStripButton19.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    toolStripMenuItem2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    toolStripMenuItem3.Checked = true;
                }

            }
            else if (toolStripButton18.Checked == false)
            {
                toolStripButton18.Checked = false;
                FontReset2();

                //斜体が有効な場合
                if (toolStripButton19.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    toolStripButton19.Checked = true;
                }

                //下線が有効な場合
                if (toolStripMenuItem2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    toolStripMenuItem2.Checked = true;
                }

                //打ち消し線が有効な場合
                if (toolStripMenuItem3.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    toolStripMenuItem3.Checked = true;
                }
            }
        }

        //斜体
        private void toolStripButton19_Click(object sender, EventArgs e)
        {
            if (kryptonRibbonGroupClusterButton2.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                    );
                }

                //完了後他のフォントスタイルを確認
                toolStripButton19.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    toolStripButton18.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    toolStripMenuItem2.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    toolStripMenuItem3.Checked = true;
                }
            }
            else if (kryptonRibbonGroupClusterButton2.Checked == false)
            {
                FontReset2();

                //太字が有効な場合
                if (toolStripButton18.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    toolStripButton18.Checked = true;
                }

                //下線が有効な場合
                if (toolStripMenuItem2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    toolStripMenuItem2.Checked = true;
                }

                //打ち消し線が有効な場合
                if (toolStripMenuItem3.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    toolStripMenuItem3.Checked = true;
                }

            }
        }

        //下線
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItem2.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                    );
                }

                //完了後他のフォントスタイルを確認
                toolStripMenuItem2.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    toolStripButton18.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    toolStripButton19.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Strikeout)
                {
                    toolStripMenuItem3.Checked = true;
                }


            }
            else if (toolStripMenuItem2.Checked == false)
            {
                FontReset2();
                //太字
                if (toolStripButton18.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    toolStripButton18.Checked = true;
                }

                //斜体
                if (toolStripButton19.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    toolStripButton19.Checked = true;
                }



                //打ち消し線
                if (toolStripMenuItem3.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                        );
                    }

                    toolStripMenuItem3.Checked = true;
                }
            }
        }

        //打ち消し線
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItem3.Checked == true)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                float fontSize;
                // 入力値がfloatに変換できるかチェック
                if (float.TryParse(kryptonRibbonGroupComboBox_NotepadFontSize.Text, out fontSize))
                {
                    // 現在のフォント名とスタイルを維持し、サイズのみ変更
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                        fontSize,
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Strikeout
                    );
                }

                //完了後他のフォントスタイルを確認
                toolStripMenuItem3.Checked = true;
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Bold)
                {
                    toolStripButton18.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Italic)
                {
                    toolStripButton19.Checked = true;
                }
                if (Notepads_kryptonRichTextBox_Notepad.SelectionFont.Underline)
                {
                    toolStripMenuItem2.Checked = true;
                }
            }
            else if (toolStripMenuItem3.Checked == false)
            {
                FontReset2();
                //太字
                if (toolStripButton18.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Bold
                        );
                    }

                    toolStripButton18.Checked = true;
                }

                //斜体
                if (toolStripButton19.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Italic
                        );
                    }

                    toolStripButton19.Checked = true;
                }

                //下線
                if (toolStripMenuItem2.Checked == true)
                {
                    Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);

                    float fontSize;
                    // 入力値がfloatに変換できるかチェック
                    if (float.TryParse(toolStripComboBox4.Text, out fontSize))
                    {
                        // 現在のフォント名とスタイルを維持し、サイズのみ変更
                        Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name,
                            fontSize,
                            Notepads_kryptonRichTextBox_Notepad.SelectionFont.Style | FontStyle.Underline
                        );
                    }

                    toolStripMenuItem2.Checked = true;
                }

            }
        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {
            using (KryptonColorDialog kryptonColorDialog = new KryptonColorDialog() { Color = Notepads_kryptonRichTextBox_Notepad.SelectionColor })
            {
                if(kryptonColorDialog.ShowDialog() ==DialogResult.OK)
                {
                    Notepads_kryptonRichTextBox_Notepad.SelectionColor = kryptonColorDialog.Color;
                    kryptonRibbonGroupColorButton2.SelectedColor = kryptonColorDialog.Color;
                }
               
            }
        }

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            using (KryptonColorDialog kryptonColorDialog = new KryptonColorDialog() { Color = Notepads_kryptonRichTextBox_Notepad.SelectionBackColor })
            {
                if (kryptonColorDialog.ShowDialog() == DialogResult.OK)
                {
                    Notepads_kryptonRichTextBox_Notepad.SelectionBackColor = kryptonColorDialog.Color;
                    kryptonRibbonGroupColorButton3.SelectedColor = kryptonColorDialog.Color;
                }

            }
        }

        private void toolStripButton22_Click(object sender, EventArgs e)
        {

        }

        private void kryptonRibbonGroupGallery2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //フォントをリセット
            Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font("メイリオ", Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size, FontStyle.Regular);
            Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
            kryptonRibbonGroupClusterButton1.Checked = false;
            kryptonRibbonGroupClusterButton2.Checked = false;
            kryptonContextMenuItem35.Checked = false;
            kryptonContextMenuItem36.Checked = false;
            toolStripButton18.Checked = false;
            toolStripButton19.Checked = false;
            toolStripMenuItem2.Checked = false;
            toolStripMenuItem3.Checked = false;
            Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
            kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;



            //選択されたアイテムに応じてフォントを変更
            if (kryptonRibbonGroupGallery2.SelectedIndex == 0)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "メイリオ",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Regular
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = false;
                kryptonRibbonGroupClusterButton2.Checked = false;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = false;
                toolStripButton19.Checked = false;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
                kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;

            }
            else if (kryptonRibbonGroupGallery2.SelectedIndex == 1)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "Yu Gothic UI Light",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Regular
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = false;
                kryptonRibbonGroupClusterButton2.Checked = false;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = false;
                toolStripButton19.Checked = false;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.FromArgb(0, 142, 197);
                kryptonRibbonGroupColorButton2.SelectedColor = Color.FromArgb(0, 142, 197);
            }
            else if (kryptonRibbonGroupGallery2.SelectedIndex == 2)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "Yu Gothic UI Light",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Regular
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = false;
                kryptonRibbonGroupClusterButton2.Checked = false;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = false;
                toolStripButton19.Checked = false;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
                kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;
            }
            else if (kryptonRibbonGroupGallery2.SelectedIndex == 3)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "Yu Gothic UI Light",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Bold
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = true;
                kryptonRibbonGroupClusterButton2.Checked = false;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = true;
                toolStripButton19.Checked = false;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
                kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;
            }
            else if (kryptonRibbonGroupGallery2.SelectedIndex == 4)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "Yu Gothic UI Light",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Italic
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = false;
                kryptonRibbonGroupClusterButton2.Checked = true;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = false;
                toolStripButton19.Checked = true;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
                kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;
            }
            else if (kryptonRibbonGroupGallery2.SelectedIndex == 5)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "メイリオ",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Bold
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = true;
                kryptonRibbonGroupClusterButton2.Checked = false;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = true;
                toolStripButton19.Checked = false;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
                kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;
            }
            else if (kryptonRibbonGroupGallery2.SelectedIndex == 6)
            {
                Notepads_kryptonRichTextBox_Notepad.Select(rtbStart, rtbLangth);
                // 現在のフォント名を変更する
                Notepads_kryptonRichTextBox_Notepad.SelectionFont = new System.Drawing.Font(
                    "メイリオ",
                    Notepads_kryptonRichTextBox_Notepad.SelectionFont.Size,
                    FontStyle.Strikeout
                );
                kryptonRibbonGroupComboBox_NotepadFont.Text = toolStripComboBox3.Text;

                kryptonRibbonGroupComboBox_NotepadFont.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();
                toolStripComboBox3.Text = Notepads_kryptonRichTextBox_Notepad.SelectionFont.Name.ToString();

                kryptonRibbonGroupClusterButton1.Checked = true;
                kryptonRibbonGroupClusterButton2.Checked = false;
                kryptonContextMenuItem35.Checked = false;
                kryptonContextMenuItem36.Checked = false;

                toolStripButton18.Checked = true;
                toolStripButton19.Checked = false;
                toolStripMenuItem2.Checked = false;
                toolStripMenuItem3.Checked = false;

                Notepads_kryptonRichTextBox_Notepad.SelectionColor = Color.Black;
                kryptonRibbonGroupColorButton2.SelectedColor = Color.Black;
            }
        }

        private void kryptonButton1_Click_1(object sender, EventArgs e)
        {
            kryptonComboBox6.Text = string.Empty;
            kryptonTextBox14.Text = string.Empty;
            kryptonTextBox8.Text = string.Empty;

        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            kryptonComboBox7.Text = string.Empty;
            kryptonTextBox9.Text = string.Empty;
            kryptonTextBox15.Text = string.Empty;
        }

        //候補1をクリックしたときの処理
        private void SearchSuggestions1_Click(object sender, EventArgs e)
        {
            if(SearchSuggestions1.Text == "印刷")
            {
            }
            else if (SearchSuggestions1.Text == "印刷プレビューを表示")
            {
            }
            else if (SearchSuggestions1.Text == "閲覧用(表示モード)")
            {

            }
            else if (SearchSuggestions1.Text == "文書作成ウィザード")
            {

            }
            else if (SearchSuggestions1.Text == "文書作成ソフトウェアで編集する")
            {

            }
            else if (SearchSuggestions1.Text == "Docxファイルとして保存")
            {

            }
            else if (SearchSuggestions1.Text == "保存(メモ帳)")
            {

            }
            else if (SearchSuggestions1.Text == "名前を付けてコピーとして保存(メモ帳)")
            {

            }
            else if (SearchSuggestions1.Text == "シート")
            {

            }
            else if (SearchSuggestions1.Text == "シート(タブ)")
            {

            }
            else if (SearchSuggestions1.Text == "シートの余白(スタイルリボンタブ)")
            {

            }
            else if (SearchSuggestions1.Text == "シートの拡大率ダイアログ")
            {

            }
            else if (SearchSuggestions1.Text == "連絡帳(タブ)")
            {

            }
            else if (SearchSuggestions1.Text == "連絡先を追加")
            {

            }
            else if (SearchSuggestions1.Text == "連絡先を削除")
            {

            }
            else if (SearchSuggestions1.Text == "メモ帳(タブ)")
            {

            }
            else if (SearchSuggestions1.Text == "余白をすべてリセット")
            {

            }
            else if (SearchSuggestions1.Text == "余白をすべてリセット")
            {

            }
            else if (SearchSuggestions1.Text == "表題のフォント(ホームリボンタブ)")
            {

            }
            else if (SearchSuggestions1.Text == "フォント(メモ帳リボンタブ)")
            {

            }
            else if (SearchSuggestions1.Text == "スタイル(ホームリボンタブ)")
            {

            }
            else if (SearchSuggestions1.Text == "表題のスタイル")
            {

            }
            else if (SearchSuggestions1.Text == "全画面モードの切り替え")
            {

            }
            else if (SearchSuggestions1.Text == "テンプレート選択画面")
            {

            }
            else if (SearchSuggestions1.Text == "新しいウィンドウを表示")
            {

            }
            else if (SearchSuggestions1.Text == "イマーシブリーダー")
            {

            }
            else if (SearchSuggestions1.Text == "設定")
            {

            }
            else if (SearchSuggestions1.Text == "テーマ設定ダイアログ")
            {

            }
            else if (SearchSuggestions1.Text == "文字の置換(シート)")
            {

            }
        }

        //候補2をクリックしたときの処理
        private void SearchSuggestions2_Click(object sender, EventArgs e)
        {

        }

        //候補3をクリックしたときの処理
        private void SearchSuggestions3_Click(object sender, EventArgs e)
        {

        }
        
        
    }

}