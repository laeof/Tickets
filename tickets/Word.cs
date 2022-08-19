using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;


namespace tickets
{
    public class Word
    { 
        /// <summary>
        /// документ
        /// </summary>
        private DocX _doc;
        /// <summary>
        /// шрифт
        /// </summary>
        private Font _fontFamily;
        /// <summary>
        /// размер шрифта
        /// </summary>
        private double _fontSize;
        /// <summary>
        /// расстояние между строками
        /// </summary>
        private float _lineSpacing;
        /// <summary>
        /// цвет шрифта
        /// </summary>
        private System.Drawing.Color _fontColor;
        /// <summary>
        /// конструктор
        /// </summary>
        /// <param name="path"></param>
        public Word(string path)
        {
            if (path != "")
                _doc = DocX.Create(path);
            else throw new Exception("CannotFindAFile");
        }
        /// <summary>
        /// установка шрифта
        /// </summary>
        /// <param name="font"></param>
        /// <param name="size"></param>
        /// <param name="color"></param>
        /// <param name="lineSpacing"></param>
        public void SetDefaultFont(string font, double size, System.Drawing.Color color, float lineSpacing)
        {
            _fontFamily = new Font(font);
            _fontSize = size;
            _fontColor = color;
            _lineSpacing = lineSpacing;
        }
        /// <summary>
        /// сохранение
        /// </summary>
        public void Save()
        {
            _doc.Save();
        }
        /// <summary>
        /// новая строка
        /// </summary>
        Paragraph _paragraph;
        /// <summary>
        /// добавление строки
        /// </summary>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <param name="color"></param>
        /// <param name="bold"></param>
        /// <param name="italic"></param>
        /// <param name="lineSpacing"></param>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        public void AddParagraph(string text, Alignment alignment, System.Drawing.Color color, bool bold = false, bool italic = false, float lineSpacing = 0, string font = null, double fontSize = 0)
        {
            //форматирование
            Formatting format = new Formatting();

            //установка формата
            if (font != null) format.FontFamily = new Font(font);
            else format.FontFamily = _fontFamily;
            if (fontSize != 0) format.Size = fontSize;
            else format.Size = _fontSize;
            format.FontColor = color;
            format.Bold = bold;
            format.Italic = italic;

            //установка параметров для строки
            _paragraph = _doc.InsertParagraph(text, false, format);
            _paragraph.Alignment = alignment;
            if (lineSpacing != 0) _paragraph.LineSpacingAfter = lineSpacing;
            else _paragraph.LineSpacingAfter = _lineSpacing;
        }
        /// <summary>
        /// новая страница
        /// </summary>
        public void NewStr()
        {
            _paragraph.InsertPageBreakAfterSelf();
        }

    }
}
