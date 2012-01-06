using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private const int HEDDING1SIZE = 24;
        private const int HEDDING2SIZE = 18;
        private const int HEDDING3SIZE = 14;

        public Form1()
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "RichTextFormat(.rtf)|*.rtf";
            if (DialogResult.OK == openFileDialog1.ShowDialog())
            {
                richTextBox1.LoadFile(openFileDialog1.FileName);
                startConvert();
            }
        }

        private void startConvert()
        {
            initializeConvert();  
            Form2 f2 = new Form2();
            f2.Show(this);
            System.Windows.Forms.Clipboard.SetText(getMessage(f2));
            f2.Dispose();
            MessageBox.Show("クリップボードにWikiフォーマットのデータがコピーされました。\r\n" +
                            "貼り付けて利用してください。",
                            "変換完了");
        }

        private string getMessage(Form2 f)
        {
            string strBuf = "";
            f.SetProgressMax(richTextBox1.TextLength);

            //目次の埋め込み
            if (checkBox1.Checked == true)
            {
                strBuf = "{{toc}}\n\n";
            }

            for (int i = 0; i < richTextBox1.TextLength; ++i)
            {
                f.SetProgressBar(1);
                strBuf += convertStyleStr(i);
            }

            //変換ミスなどの削除
            strBuf = strBuf.Replace("** \n", "\n");
            strBuf = strBuf.Replace("__ \n", "\n");

            return strBuf;
        }

        int   m_iFontSize       = 9;
        bool  m_bIsColorChanged = false;
        bool  m_bIsBold         = false;
        bool  m_bIsItalic       = false;
        bool  m_bIsChangePos    = false;
        bool  m_bIsHedding      = false;
        Color m_colBeforeColor  = Color.Black;

        private void initializeConvert()
        {
            m_bIsColorChanged = false;
            m_colBeforeColor = Color.Black;
            m_bIsBold = false;
            m_bIsItalic = false;
            m_iFontSize = 9;
        }

        private string convertStyleStr(int index)
        {
            richTextBox1.Select(index, 1);
            string strbuf = "";

            int iFontSize = (int)richTextBox1.SelectionFont.Size;
            if (m_iFontSize != iFontSize)
            {
                //見出し１
                if (iFontSize >= HEDDING1SIZE && m_bIsHedding == false)
                {
                    strbuf += "h1. ";
                    m_bIsHedding = true;
                }
                //見出し２
                else if (iFontSize >= HEDDING2SIZE && m_bIsHedding == false)
                {
                    strbuf += "h2. ";
                    m_bIsHedding = true;
                }
                //見出し３
                else if (iFontSize >= HEDDING3SIZE && m_bIsHedding == false)
                {
                    strbuf += "h3. ";
                    m_bIsHedding = true;
                }
            }

            {  //ボールドの開始と終了
                if (iFontSize < HEDDING3SIZE && m_bIsBold == false && richTextBox1.SelectionFont.Bold)
                {
                    m_bIsBold = true;
                    strbuf += "*";
                }

                if (m_bIsBold == true && !richTextBox1.SelectionFont.Bold)
                {
                    m_bIsBold = false;
                    strbuf += "* ";
                }
            }

            {  //イタリックの開始と終了
                if (m_bIsItalic == false && richTextBox1.SelectionFont.Italic)
                {
                    m_bIsItalic = true;
                    strbuf += " _";
                }

                if (m_bIsItalic == true && !richTextBox1.SelectionFont.Italic)
                {
                    m_bIsItalic = false;
                    strbuf += "_ ";
                }
            }

            // センタリング
            if (m_bIsChangePos == false && richTextBox1.SelectionAlignment == HorizontalAlignment.Center)
            {
                strbuf += "p=. ";
                m_bIsChangePos = true;
            }

            // 右寄せ
            if (m_bIsChangePos == false && richTextBox1.SelectionAlignment == HorizontalAlignment.Right)
            {
                strbuf += "p>. ";
                m_bIsChangePos = true;
            }

            // 色の変更開始と終了
            if (m_colBeforeColor != richTextBox1.SelectionColor && richTextBox1.SelectedText != "\n")
            {
                m_colBeforeColor = richTextBox1.SelectionColor;

                if (m_bIsColorChanged == true)
                {
                    strbuf += "% ";
                    m_bIsColorChanged = false;
                }

                if (richTextBox1.SelectionColor != Color.Black)
                {
                    string colorname = ColorTranslator.ToHtml(richTextBox1.SelectionColor);
                    strbuf += " %{color:" + colorname + "}";
                    m_bIsColorChanged = true;
                }
            }


            //改行前の処理
            if (richTextBox1.SelectedText == "\n")
            {
                if (m_bIsBold == true)
                {
                    strbuf += "* ";
                    m_bIsBold = false;
                }

                if (m_bIsItalic == true)
                {
                    strbuf += "_ ";
                    m_bIsItalic = false;
                }

                if (m_colBeforeColor != Color.Black)
                {
                    strbuf += "%";
                    m_colBeforeColor = Color.Black;
                    m_bIsColorChanged = false;
                }

                if (m_bIsHedding == true) strbuf += "\n";
                m_bIsHedding = false;

                m_bIsChangePos = false;
            }

            //文字のコピー
            strbuf += richTextBox1.SelectedText;

            //テーブルの変換
            strbuf = strbuf.Replace("●", "* ");
            strbuf = strbuf.Replace("○", "** ");
            strbuf = strbuf.Replace("■", "*** ");


            return strbuf;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("読み込みボタンを押してリッチテキスト形式のファイルを読み込みます\r\n" +
                            "読み込みが完了したら自動的に変換されます。\r\n" +
                            "\r\n" +
                            "変換が完了するとダイアログメッセージが表示され\r\n" +
                            "クリップボードに変換後のデータがコピーされます。\r\n" + 
                            "直接wikiに張り付けてチェックしてください。\r\n" +
                            "\r\n" +
                            "また、現在表示されているテキストを出力することができます。\r\n" +
                            "表示中の文章を変換ボタンを押すと出力されます。\r\n" +
                            "初期化ボタンを押すと、文章やスタイルが初期化されます。\r\n" +
                            "\r\n" +
                            "正しく変換されない場合がありますがご了承ください。",
                            "使い方",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            startConvert();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            checkBox1.Checked = true;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }
    }
}
