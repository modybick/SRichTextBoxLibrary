using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Printing;

namespace SRichTextBoxLibrary
{
    /// <summary>
    /// RichTextBoxを拡張したクラス
    /// </summary>
    public partial class SuperRichTextBox: RichTextBox
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public SuperRichTextBox()
        {
            InitializeComponent();
            this.LanguageOption = RichTextBoxLanguageOptions.UIFonts;
        }

        /// <summary>
        /// ショートカットキーの設定
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">KeyEventArgs</param>
        private void SRichTextBox_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.B && e.Control && !e.Shift)
            {   //Bold  [ctrl+B]
                toggleFontStyle(FontStyle.Bold);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.U && e.Control && !e.Shift)
            {   //Underline [ctrl+U]
                toggleFontStyle(FontStyle.Underline);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.T && e.Control && !e.Shift)
            {   //Strikeout [ctrl+T]
                toggleFontStyle(FontStyle.Strikeout);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.I && e.Control && !e.Shift)
            {   //Itaric [ctrl+I]
                toggleFontStyle(FontStyle.Italic);
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.F && e.Control && !e.Shift)
            {   //フォントダイアログ
                showFontDialog();
                e.SuppressKeyPress = true;
            }

            if (e.KeyCode == Keys.A && e.Control && !e.Shift)
            {   //全選択 [ctrl+A]
                SelectAll();
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// フォントスタイルの切り替え（トグル）
        /// </summary>
        /// <param name="style">ON/OFFするフォントスタイル</param>
        private void toggleFontStyle(FontStyle style)
        {
            if (SelectionLength == 0)
            {   //選択されていない場合
                SelectionFont = new Font(SelectionFont, SelectionFont.Style ^ style);
                return;
            }
            //作業用RichTextBoxを生成
            RichTextBox bufRtb = new RichTextBox();
            //bufRtbにテキスト貼り付け
            bufRtb.Rtf = this.Rtf;

            int selectionStart = this.SelectionStart;   //選択開始位置
            int selectionLength = this.SelectionLength; //選択範囲の長さ
            int selectionEnd = selectionStart + selectionLength;    //選択終了位置

            bufRtb.Select(selectionStart, 1);
            if (bufRtb.SelectionFont.Style >= style)
            {   //一文字目が指定のスタイルを含む場合
                for (int x = selectionStart; x < selectionEnd; ++x)
                {   //一文字目から終了位置までループ
                    bufRtb.Select(x, 1);  //一文字ずつ選択
                    if (bufRtb.SelectionFont.Style >= style)
                    {
                        bufRtb.SelectionFont =
                            new Font(bufRtb.SelectionFont, bufRtb.SelectionFont.Style ^ style);
                    }
                }
            } else
            {   //一文字目が指定のスタイルを含まない場合
                for (int x = selectionStart; x < selectionEnd; ++x)
                {   //一文字目から終了位置までループ
                    bufRtb.Select(x, 1);  //一文字ずつ選択
                    bufRtb.SelectionFont =
                        new Font(bufRtb.SelectionFont, bufRtb.SelectionFont.Style | style);
                }
            }

            bufRtb.Select(selectionStart, selectionLength);
            if (this.SelectedRtf.EndsWith(""))
            {   //選択範囲の最後が空白で終わる場合
                this.Select(SelectionStart, selectionLength);
            }
            this.SelectedRtf = bufRtb.SelectedRtf;
            bufRtb.Dispose();
            //元の選択に戻す
            this.Select(selectionStart, selectionLength);
        }

        /// <summary>
        /// フォントダイアログの表示
        /// </summary>
        private void showFontDialog()
        {
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                this.SelectionFont = fontDialog1.Font;
            }
        }

        /********************************************************
        ** ContextMenuStrip
        */

        /// <summary>
        /// メニューストリップが開いた時
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">CancelEventArgs</param>
        private void contextMenuStrip1_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {   
            if (this.SelectionLength <= 0)
            {   //何も選択されていなければ、切り取り・コピーは選べない   
                this.cutMenuItem.Enabled = false;
                this.copyMenuItem.Enabled = false;
            } else
            {   //何か選択されていれば、切り取り・コピーは選択可能
                this.cutMenuItem.Enabled = true;
                this.copyMenuItem.Enabled = true;
            }
        }

        /// <summary>
        /// メニュー_切り取り
        /// </summary>
        private void cutMenuItem_Click(object sender, EventArgs e)
        {
            this.Cut();
        }

        /// <summary>
        /// メニュー_コピー
        /// </summary>
        private void copyMenuItem_Click(object sender, EventArgs e)
        {
            this.Copy();
        }

        /// <summary>
        /// メニュー_貼り付け
        /// </summary>
        private void pasteMenuItem_Click(object sender, EventArgs e)
        {
            this.Paste();
        }

        /// <summary>
        /// メニュー_フォントダイアログ
        /// </summary>
        private void fontDialogMenuItem_Click(object sender, EventArgs e)
        {
            showFontDialog();
        }

        /*******************************************************
         * 印刷関係
         */

        /// <summary>
        /// プリントドキュメントを取得する。
        /// </summary>
        /// <returns>PrintDocumentを返す。</returns>
        public PrintDocument getPrintDocument()
        {
            return this.printDocument1;
        }

        /// <summary>
        /// プリント確認
        /// </summary>
        private int checkPrint;

        /// <summary>
        /// 印刷開始時
        /// </summary>
        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            checkPrint = 0;
        }

        /// <summary>
        /// 印刷時（各ページ）
        /// </summary>
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            // Print the content of RichTextBox. Store the last character printed.
            checkPrint = this.Print(checkPrint, this.TextLength, e);

            // Check for more pages
            if (checkPrint < this.TextLength)
                e.HasMorePages = true;
            else
                e.HasMorePages = false;
        }
        
        /// <summary>
        /// 印刷
        /// </summary>
        public void print()
        {
            if (printDialog1.ShowDialog() == DialogResult.OK)
                printDocument1.Print();
        }

        /// <summary>
        /// 印刷プレビュー
        /// </summary>
        public void showPrintPreview()
        {
            printPreviewDialog1.ShowDialog();
        }

        /// <summary>
        /// 印刷ページ設定
        /// </summary>
        public void printPageSetup()
        {
            pageSetupDialog1.ShowDialog();
        }

        //Convert the unit used by the .NET framework (1/100 inch) 
        //and the unit used by Win32 API calls (twips 1/1440 inch)
        private const double anInch = 14.4;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct CHARRANGE
        {
            public int cpMin;         //First character of range (0 for start of doc)
            public int cpMax;           //Last character of range (-1 for end of doc)
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FORMATRANGE
        {
            public IntPtr hdc;             //Actual DC to draw on
            public IntPtr hdcTarget;       //Target DC for determining text formatting
            public RECT rc;                //Region of the DC to draw to (in twips)
            public RECT rcPage;            //Region of the whole DC (page size) (in twips)
            public CHARRANGE chrg;         //Range of text to draw (see earlier declaration)
        }

        private const int WM_USER = 0x0400;
        private const int EM_FORMATRANGE = WM_USER + 57;

        [DllImport("USER32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, IntPtr lp);

        // Render the contents of the RichTextBox for printing
        //	Return the last character printed + 1 (printing start from this point for next page)
        public int Print(int charFrom, int charTo, PrintPageEventArgs e)
        {
            //Calculate the area to render and print
            RECT rectToPrint;
            rectToPrint.Top = (int)(e.MarginBounds.Top * anInch);
            rectToPrint.Bottom = (int)(e.MarginBounds.Bottom * anInch);
            rectToPrint.Left = (int)(e.MarginBounds.Left * anInch);
            rectToPrint.Right = (int)(e.MarginBounds.Right * anInch);

            //Calculate the size of the page
            RECT rectPage;
            rectPage.Top = (int)(e.PageBounds.Top * anInch);
            rectPage.Bottom = (int)(e.PageBounds.Bottom * anInch);
            rectPage.Left = (int)(e.PageBounds.Left * anInch);
            rectPage.Right = (int)(e.PageBounds.Right * anInch);

            IntPtr hdc = e.Graphics.GetHdc();

            FORMATRANGE fmtRange;
            fmtRange.chrg.cpMax = charTo;               //Indicate character from to character to 
            fmtRange.chrg.cpMin = charFrom;
            fmtRange.hdc = hdc;                    //Use the same DC for measuring and rendering
            fmtRange.hdcTarget = hdc;              //Point at printer hDC
            fmtRange.rc = rectToPrint;             //Indicate the area on page to print
            fmtRange.rcPage = rectPage;            //Indicate size of page

            IntPtr res = IntPtr.Zero;

            IntPtr wparam = IntPtr.Zero;
            wparam = new IntPtr(1);

            //Get the pointer to the FORMATRANGE structure in memory
            IntPtr lparam = IntPtr.Zero;
            lparam = Marshal.AllocCoTaskMem(Marshal.SizeOf(fmtRange));
            Marshal.StructureToPtr(fmtRange, lparam, false);

            //Send the rendered data for printing 
            res = SendMessage(Handle, EM_FORMATRANGE, wparam, lparam);

            //Free the block of memory allocated
            Marshal.FreeCoTaskMem(lparam);

            //Release the device context handle obtained by a previous call
            e.Graphics.ReleaseHdc(hdc);

            //Return last + 1 character printer
            return res.ToInt32();
        }


    }

}
