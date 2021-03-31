using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;


namespace comtray
{
    /// <summary>
    /// This is a CODE to FONT and COLOR mapping entry.
    /// </summary>
    /// 
    /// Used to select RED, GRN or Normal text on the richtext box.
    public class FontCodePair
    {
        public int code;
        public Font font;
        public Color color;
        public FontCodePair( int code, Font f, Color c )
        {
            this.font = f;
            this.color = c;
            this.code = code;
        }
    }

    /// <summary>
    /// This is a Modified RichTextBox that accepts CODED text
    /// </summary>
    public class CodedRichTextBox : RichTextBox
    {
        private const short WM_PAINT = 0x00f;
        private List<FontCodePair> fontpairs = new List<FontCodePair>();


        public bool _Paint = true;

        public CodedRichTextBox()
        {
            this.Location = new Point(2, 2);
            this.Dock = DockStyle.Fill;
            this.WordWrap = false;
            this.ReadOnly = true;
            Color c = this.ForeColor;
            Font f;

            // we only have 3 codes, Normal, RED and GREEN
            f = new Font("Courier New", this.Font.Size);
            fontpairs.Add(new FontCodePair(0, f, c));
            c = Color.FromName("green");
            f = new Font("Courier New", this.Font.Size, FontStyle.Bold);
            fontpairs.Add(new FontCodePair(+1, f, c));

            c = Color.FromName("red");
            f = new Font("Courier New", this.Font.Size, FontStyle.Bold | FontStyle.Strikeout);
            fontpairs.Add(new FontCodePair(-1, f, c));

            this.ScrollBars = RichTextBoxScrollBars.Both;
            this.ShortcutsEnabled = true;
            this.ZoomFactor = 1;
        }

        /// Adds coded text to the RichTextBox per the colorizing code.
        public void CodedText(string text, int code)
        {
            int start;
            
            start = this.TextLength;

            this.AppendText(text);

            foreach (FontCodePair tmp in fontpairs)
            {
                if (code == tmp.code)
                {
                    this.Select(start, text.Length);
                    this.SelectionColor = tmp.color;
                    this.SelectionFont = tmp.font;
                    return;
                }
            }
            /* default, do nothing */
            throw new NotImplementedException("Missing CODE");
        }
    }
}
