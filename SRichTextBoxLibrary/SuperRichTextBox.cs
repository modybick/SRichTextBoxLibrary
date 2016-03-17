using System.Drawing;
using System.Windows.Forms;

namespace SRichTextBoxLibrary
{
    public partial class SuperRichTextBox: RichTextBox
    {
        public SuperRichTextBox()
        {
            InitializeComponent();
            this.LanguageOption = RichTextBoxLanguageOptions.UIFonts;
        }

        private void SRichTextBox_KeyDown(object sender, KeyEventArgs e)
        {   //ショートカットの設定

            if (e.KeyCode == Keys.B && e.Control)
            {   //Bold  [ctrl+B]
                toggleFontStyle(FontStyle.Bold);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.U && e.Control)
            {   //Underline [ctrl+U]
                toggleFontStyle(FontStyle.Underline);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.T && e.Control)
            {   //Strikeout [ctrl+T]
                toggleFontStyle(FontStyle.Strikeout);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.I && e.Control)
            {   //Itaric [ctrl+I]
                toggleFontStyle(FontStyle.Italic);
                e.SuppressKeyPress = true;
            }

        }

        private void toggleFontStyle(FontStyle style)
        {   //フォントスタイルを切り替え（トグル）

            //作業用RichTextBoxを生成
            RichTextBox bufRtb = new RichTextBox();
            //bufRtbにテキスト貼り付け
            bufRtb.Rtf = this.SelectedRtf;

            int selectionStart = this.SelectionStart;   //選択開始位置
            int selectionLength = this.SelectionLength; //選択範囲の長さ
            int selectionEnd = selectionStart + selectionLength;    //選択終了位置

            for ( int x = 0; x < selectionLength; ++x )
            {   //選択開始位置から終了位置までループ
                bufRtb.Select(x, 1);  //一文字ずつ選択
                if (this.SelectionFont.Style >= style)
                {   //一文字目が指定のスタイルを含む場合
                    bufRtb.SelectionFont =
                        new Font(bufRtb.SelectionFont, bufRtb.SelectionFont.Style ^ style);

                }
                else
                {   //一文字目が指定のスタイルを含まない場合
                    bufRtb.SelectionFont =
                        new Font(bufRtb.SelectionFont, bufRtb.SelectionFont.Style | style);

                }

            }
            bufRtb.Select(0, selectionLength);
            this.SelectedRtf = bufRtb.SelectedRtf;
            bufRtb.Dispose();
            //元の選択に戻す
            this.Select(selectionStart, selectionLength);
        }

    }

}
