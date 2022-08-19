using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace tickets
{
    public class PdfWrite
    {
        /// <summary>
        /// шрифт
        /// </summary>
        private string ttf = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "Arial.ttf");
        /// <summary>
        /// документ
        /// </summary>
        private Document document;
        /// <summary>
        /// запись
        /// </summary>
        private PdfWriter writer;
        /// <summary>
        /// базовый шрифт
        /// </summary>
        private BaseFont baseFont;
        /// <summary>
        /// шрифт 1
        /// </summary>
        private Font font;
        /// <summary>
        /// шрифт 2
        /// </summary>
        private Font h_font;
        /// <summary>
        /// шрифт 3
        /// </summary>
        private Font t_h_font;
        /// <summary>
        /// конструктор
        /// </summary>
        /// <param name="_filename"></param>
        /// <param name="_l"></param>
        /// <param name="_r"></param>
        /// <param name="_t"></param>
        /// <param name="_b"></param>
        public PdfWrite(string _filename, float _l = 40, float _r = 40, float _t = 30, float _b = 50)
        {
            //документ
            document = new Document(PageSize.A4, _l, _r, _t, _b);
            //запись
            if (_filename != "")
                writer = PdfWriter.GetInstance(document,
                    new FileStream(_filename + ".pdf", FileMode.Create));
            else throw new Exception("CannotFindAFile");
            _filename += ".pdf";
            //відкриваємо
            document.Open();
            document.NewPage();

            //шрифти
            baseFont = BaseFont.CreateFont(ttf, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            t_h_font = new iTextSharp.text.Font(baseFont, 12f, iTextSharp.text.Font.UNDERLINE);
            h_font = new iTextSharp.text.Font(baseFont, 12f, iTextSharp.text.Font.BOLD);
            font = new iTextSharp.text.Font(baseFont, 12f, iTextSharp.text.Font.NORMAL);
        }
        /// <summary>
        /// добавляємо заголовок
        /// </summary>
        /// <param name="_text"></param>
        /// <param name="_aligment"></param>
        /// <param name="_s"></param>
        /// <param name="_font"></param>
        /// <param name="_o"></param>
        public void AddHeader(string _text, int _aligment, float _s, int _font, float _o = 1.16f)
        {
            Paragraph h;
            switch (_font)
            {
                case 1:
                    //under
                    h = new Paragraph(_text, t_h_font);
                    h.Alignment = _aligment;
                    h.SetLeading(0.0f, _o);
                    h.SpacingAfter = _s;
                    document.Add(h);
                    break;
                case 2:
                    //bold
                    h = new Paragraph(_text, h_font);
                    h.Alignment = _aligment;
                    h.SetLeading(0.0f, _o);
                    h.SpacingAfter = _s;
                    document.Add(h);
                    break;
                case 3:
                    //normal
                    h = new Paragraph(_text, font);
                    h.Alignment = _aligment;
                    h.SetLeading(0.0f, _o);
                    h.SpacingAfter = _s;
                    document.Add(h);
                    break;
            }

        }
        /// <summary>
        /// нова сторінка
        /// </summary>
        public void Newlist()
        {
            document.NewPage();
        }
        /// <summary>
        /// закриваємо документ
        /// </summary>
        public void Write()
        {
            document.Close();
            writer.Close();
        }
    }
}
