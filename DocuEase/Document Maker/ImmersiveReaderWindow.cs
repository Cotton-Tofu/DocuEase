using FluentTransitions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace Document_Maker
{
    
    public partial class ImmersiveReaderWindow : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        //Rtfファイルを読み込むかシートの内容を読むかのフラグ
        public bool IsRtfRead { get; set; }

        //Form1のシートの内容を取得する
        //文書入力用
        //発行番号
        public string IssueNumber{ get; set; }
        //日付
        public string Date { get; set; }
        //発行番号
        //宛先用
        public string AdCompany { get; set; }
        public string AdTitle { get; set; }
        public string AdName { get; set; }
        //発信者用
        public string CaCampany { get; set; }
        public string CaLocation { get; set; }
        public string CaBuildingName { get; set; }
        public string CaTitle { get; set; }
        public string CaName { get; set; }
        public string CaMailAddress { get; set; }
        public string CaPhoneNumber1 { get; set; }
        public string CaFaxNumber1 { get; set; }
        //表題
        public string title { get; set; }
        //あいさつ文
        public string Greeting { get; set; }
        //感謝のあいさつ
        public string ThankYouGreeting { get; set; }
        //結語
        public string Conclusion { get; set; }
        //内容
        public string Content { get; set; }
        //記
        public string Note { get; set; }
        //記し書き
        public string Notetaking { get; set; }
        //以上
        public string Notetaking_End { get; set; }


        SpeechSynthesizer voice = new SpeechSynthesizer();
        string originalRtf = null;
       

        public ImmersiveReaderWindow()
        {
            InitializeComponent();
            voice.SpeakCompleted += Voice_SpeakCompleted;
            voice.SpeakProgress += Voice_SpeekProgress;

            timer.Tick += timer_Tick;
            timer.Interval = 1000;
        }

        private void Voice_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            if (kryptonCheckButton1.Text == "再生中")
            {
                kryptonRichTextBox1.SelectionFont = new Font("メイリオ", 18, FontStyle.Regular);

                kryptonCheckButton1.Checked = false;
                kryptonCheckButton1.Text = "一時停止中";

                voice.Pause();

                timer.Stop();
                Check60Seconds = 0;
                CheckOneMinuts = 0;

                toolStripStatusLabel5.Text = "読み終わりました";

                // 読み終わったら元の RTF を復元
                if (!string.IsNullOrEmpty(originalRtf))
                {
                    kryptonRichTextBox1.Rtf = originalRtf;
                    originalRtf = null;
                    kryptonRichTextBox1.Select(0, 0);
                }
            }
        }

        private void Voice_SpeekProgress(object sender, SpeakProgressEventArgs e)
        {
            if (kryptonRichTextBox1.IsDisposed) return;

            // イベントはワーカースレッドで来るので UI スレッドへ移す
            BeginInvoke((Action)(() =>
            {
                try
                {
                    // 毎回元の RTF に戻してから、現在の発話部分だけ太字にする
                    if (!string.IsNullOrEmpty(originalRtf))
                        kryptonRichTextBox1.Rtf = originalRtf;

                    int pos = e.CharacterPosition;
                    int len = e.Text?.Length ?? 0;

                    if (len > 0 && pos >= 0 && pos + len <= kryptonRichTextBox1.TextLength)
                    {
                        kryptonRichTextBox1.Select(pos, len);
                        var selFont = kryptonRichTextBox1.SelectionFont ?? kryptonRichTextBox1.Font;
                        kryptonRichTextBox1.SelectionFont = new Font(selFont, selFont.Style | FontStyle.Bold);
                    }
                    else
                    {
                        kryptonRichTextBox1.SelectionLength = 0;
                    }

                }
                catch
                {
                    // UI 更新中に失敗しても落とさない
                }
            }));
        }

        private void kryptonCheckButton1_Click(object sender, EventArgs e)
        {
            try
            {
                if (kryptonCheckButton1.Text == "再生中")
                {
                    kryptonCheckButton1.Checked = false;
                    kryptonCheckButton1.Text = "一時停止中";

                    voice.Pause();
                    timer.Stop();

                    toolStripStatusLabel5.Text = "一時停止";
                }
                else if (kryptonCheckButton1.Text == "一時停止中")
                {
                    kryptonRichTextBox1.SelectionFont = new Font("メイリオ", 18 , FontStyle.Regular);
                    kryptonCheckButton1.Checked = true;
                    kryptonCheckButton1.Text = "再生中";

                    if (kryptonComboBox1.SelectedIndex == 0)
                    {
                        voice.SelectVoiceByHints(VoiceGender.NotSet);
                    }
                    else if (kryptonComboBox1.SelectedIndex == 1)
                    {
                        voice.SelectVoiceByHints(VoiceGender.Female);
                    }
                    else if (kryptonComboBox1.SelectedIndex == 2)
                    {
                        voice.SelectVoiceByHints(VoiceGender.Male);
                    }
                    else if (kryptonComboBox1.SelectedIndex == 3)
                    {
                        voice.SelectVoiceByHints(VoiceGender.Neutral);
                    }

                    // 再生開始直前に元の RTF を保存
                    originalRtf = kryptonRichTextBox1.Rtf;

                    voice.Resume();
                    voice.SpeakAsync(kryptonRichTextBox1.Text);
                    voice.Volume = kryptonTrackBar1.Value;

                    timer.Start();

                    toolStripStatusLabel5.Text = "スピーチ中";
                }
            }
            catch(Exception ex) 
            {
                MessageBox.Show("スピーチを実行することができませんでした。詳細は以下を確認してください。\r\n"+ex.Message,"実行失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            try
            {
                //保存用のSpeechSynthesizerインスタンスを作成
                SpeechSynthesizer SaveSS = new SpeechSynthesizer();
                using (SaveFileDialog sd = new SaveFileDialog() { Title = "音声ファイルを保存する場所を選択",Filter = "WAVファイル(*.wav)|:.wav" })
                {
                    if(sd.ShowDialog() == DialogResult.OK)
                    {
                        //音声設定
                        if (kryptonComboBox1.SelectedIndex == 0)
                        {
                            SaveSS.SelectVoiceByHints(VoiceGender.NotSet);
                        }
                        else if (kryptonComboBox1.SelectedIndex == 1)
                        {
                            SaveSS.SelectVoiceByHints(VoiceGender.Female);
                        }
                        else if (kryptonComboBox1.SelectedIndex == 2)
                        {
                            SaveSS.SelectVoiceByHints(VoiceGender.Male);
                        }
                        else if (kryptonComboBox1.SelectedIndex == 3)
                        {
                            SaveSS.SelectVoiceByHints(VoiceGender.Neutral);
                        }
                        //音量設定
                        SaveSS.Volume = kryptonTrackBar1.Value;

                        //保存
                        SaveSS.SetOutputToWaveFile(sd.FileName);
                        SaveSS.Speak(kryptonRichTextBox1.Text);
                        SaveSS.SetOutputToDefaultAudioDevice();

                        kryptonLabel3.Text = voice.Volume.ToString() + "％";

                        //保存が完了したら破棄する
                        SaveSS.Dispose();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("WAVファイルとして保存することができませんでした。詳細は以下を確認してください。\r\n" + ex.Message, "保存失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void kryptonTrackBar1_ValueChanged(object sender, EventArgs e)
        {
            voice.Volume = kryptonTrackBar1.Value;

            kryptonLabel3.Text = voice.Volume.ToString() + "％";
        }
        
        Timer timer = new Timer();


        private void kryptonCheckButton2_Click(object sender, EventArgs e)
        {
            if(kryptonCheckButton2.Checked == true)
            {
                kryptonTrackBar1.Enabled = false;
                voice.Volume = 0;

                kryptonLabel3.Text = voice.Volume.ToString() + "％";
            }
            else
            {
                kryptonTrackBar1.Enabled = true;
                voice.Volume = kryptonTrackBar1.Value;

                kryptonLabel3.Text = voice.Volume.ToString() + "％";
            }
        }


        int Check60Seconds = 0;
        int CheckOneMinuts = 0;
        private void timer_Tick(object sender, EventArgs e)
        {
               

            Check60Seconds = Check60Seconds + 1;

            if (Check60Seconds == 0)
            {
                toolStripStatusLabel3.Text = "00";
            }
            else if (Check60Seconds == 1)
            {
                toolStripStatusLabel3.Text = "01";
            }
            else if (Check60Seconds == 2)
            {
                toolStripStatusLabel3.Text = "02";
            }
            else if (Check60Seconds == 3)
            {
                toolStripStatusLabel3.Text = "03";
            }
            else if (Check60Seconds == 4)
            {
                toolStripStatusLabel3.Text = "04";
            }
            else if (Check60Seconds == 5)
            {
                toolStripStatusLabel3.Text = "05";
            }
            else if (Check60Seconds == 6)
            {
                toolStripStatusLabel3.Text = "06";
            }
            else if (Check60Seconds == 7)
            {
                toolStripStatusLabel3.Text = "07";
            }
            else if (Check60Seconds == 8)
            {
                toolStripStatusLabel3.Text = "08";
            }
            else if (Check60Seconds == 9)
            {
                toolStripStatusLabel3.Text = "09";
            }
            else
            {
                toolStripStatusLabel3.Text = Check60Seconds.ToString();
                if(Check60Seconds == 60)
                {
                    Check60Seconds = 0;
                    toolStripStatusLabel3.Text = "00";
                    CheckOneMinuts = CheckOneMinuts + 1;
                    if (CheckOneMinuts == 0)
                    {
                        toolStripStatusLabel1.Text = "00";
                    }
                    else  if (CheckOneMinuts == 1)
                    {
                        toolStripStatusLabel1.Text = "01";
                    }
                    else if (CheckOneMinuts == 2)
                    {
                        toolStripStatusLabel1.Text = "02";
                    }
                    else if (CheckOneMinuts == 3)
                    {
                        toolStripStatusLabel1.Text = "03";
                    }
                    else if (CheckOneMinuts == 4)
                    {
                        toolStripStatusLabel1.Text = "04";
                    }
                    else if (CheckOneMinuts == 5)
                    {
                        toolStripStatusLabel1.Text = "05";
                    }
                    else if (CheckOneMinuts == 6)
                    {
                        toolStripStatusLabel1.Text = "06";
                    }
                    else if (CheckOneMinuts == 7)
                    {
                        toolStripStatusLabel1.Text = "07";
                    }
                    else if (CheckOneMinuts == 8)
                    {
                        toolStripStatusLabel1.Text = "08";
                    }
                    else if (CheckOneMinuts == 9)
                    {
                        toolStripStatusLabel1.Text = "09";
                    }
                    else
                    {
                        toolStripStatusLabel1.Text = CheckOneMinuts.ToString();
                    }

                }
            }



        }

        private void kryptonCheckButton3_Click(object sender, EventArgs e)
        {
            if(kryptonCheckButton3.Checked == true)
            {
                this.WindowState = FormWindowState.Normal;
                this.AllowFormChrome = false;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.AllowFormChrome = true;
                this.FormBorderStyle = FormBorderStyle.Sizable;
                this.WindowState = FormWindowState.Normal;
            }

        }

        private void kryptonCheckButton2_Click_1(object sender, EventArgs e)
        {

        }

        private void kryptonPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            kryptonRichTextBox1.SelectionFont = new Font("メイリオ", 18, FontStyle.Regular);

            kryptonCheckButton1.Checked = false;
            kryptonCheckButton1.Text = "一時停止中";

            voice.Pause();
            voice.SpeakAsyncCancelAll();
            timer.Stop();

            // キャンセル時に元の RTF を復元
            if (!string.IsNullOrEmpty(originalRtf))
            {
                kryptonRichTextBox1.Rtf = originalRtf;
                originalRtf = null;
                kryptonRichTextBox1.Select(0, 0);
            }

            Check60Seconds = 0;
            CheckOneMinuts = 0;
            toolStripStatusLabel1.Text = "00";
            toolStripStatusLabel3.Text = "00";


            toolStripStatusLabel5.Text = "準備完了";
        }

        private void ImmersiveReaderWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            voice.Pause();
            voice.SpeakAsyncCancelAll();
            timer.Stop();

            // 終了時に元に戻す
            if (!string.IsNullOrEmpty(originalRtf))
            {
                kryptonRichTextBox1.Rtf = originalRtf;
                originalRtf = null;
            }

            voice.Dispose();
            timer.Dispose();
        }

        private void kryptonRichTextBox1_VScroll(object sender, EventArgs e)
        {
        }

        private void ImmersiveReaderWindow_Load(object sender, EventArgs e)
        {
            kryptonRichTextBox1.Clear();
            if (IsRtfRead == false)
            {
                //シートの内容をリッチテキストエディタに表示する
                //string.Emptyは使えないのでnullで代用する
                if (IssueNumber != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += IssueNumber;
                }
                if (Date != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Date;
                }
                if (AdCompany != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += AdCompany;
                }
                if (AdTitle != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += AdTitle;
                }
                if (AdName != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += AdName;
                }
                if (CaCampany != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += CaCampany;
                }
                if (CaLocation != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += CaLocation;
                }
                if (CaBuildingName != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += CaBuildingName;
                }
                if (CaTitle != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += CaTitle;
                }
                if (CaName != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += CaName;
                }
                if (CaMailAddress != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += "メールアドレス:" + CaMailAddress;
                }
                if (CaPhoneNumber1 != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += "電話番号:" + CaPhoneNumber1;
                }
                if (CaFaxNumber1 != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += "Fax番号:" + CaFaxNumber1;
                }
                if (title != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += title;
                }
                if (Greeting != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Greeting;
                }
                if (ThankYouGreeting != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += ThankYouGreeting;
                }
                if (Content != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Content;
                }
                if (Conclusion != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Conclusion;
                }
                if (Note != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Note;
                }
                if (Notetaking != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Notetaking;
                }
                if (Notetaking != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Notetaking;
                }
                if (Notetaking_End != null)
                {
                    kryptonRichTextBox1.Text += "\n";
                    kryptonRichTextBox1.Text += Notetaking_End;
                }
            }
            else if (IsRtfRead == true)
            {
                try
                {
                    //メモの内容を復元
                    String str = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\DocuEase\SaveFile.rtf";

                    if (File.Exists(str))
                    {
                        kryptonRichTextBox1.LoadFile(str);
                    }
                    else
                    {
                        MessageBox.Show("自動保存ファイルがないためイマーシブリーダーを表示できません。", "イマーシブリーダー", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        this.Close();
                    }
                }
                catch
                { }
            }


            //テーマ設定
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

        private void ImmersiveReaderWindow_Shown(object sender, EventArgs e)
        {

        }

        private void ImmersiveReaderWindow_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void kryptonRichTextBox1_SelectionChanged(object sender, EventArgs e)
        {
            
        }
    }
}
