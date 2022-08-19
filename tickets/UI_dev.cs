using System;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using Xceed.Words.NET;

namespace tickets
{
    public partial class UI_dev : Form
    {
        public int NUMBERSOFTICKETS = 25;
        public List<string> s = new List<string>();
        public UI_dev()
        {
            InitializeComponent();
            ticketlistopen();
        }
        /// <summary>
        /// рандомізація
        /// </summary>
        Random rnd = new Random();
        /// <summary>
        /// кидаємо у пдф
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (CheckComboboxesAreNotNull())
            {

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //екземпляр класу
                    PdfWrite p = new PdfWrite(saveFileDialog.FileName);
                    for (int i = 0; i < NUMBERSOFTICKETS; i++)
                    {
                        p.AddHeader("Національний аерокосмічний університет ім. М.Є. Жуковського 'ХАІ'", 1, 10f, 2);
                        p.AddHeader(("Курс " + comboBox1.Text + ", ") + ("cеместр " + comboBox2.Text + ", ") + ("навчальний рік " + (comboBox3.Text) + '-' + comboBox4.Text), 0, 1.16f, 3);
                        p.AddHeader("Спеціальність " + comboBox5.Text, 0, 1.16f, 3);
                        p.AddHeader("Спеціальність " + comboBox6.Text, 0, 1.16f, 3);
                        p.AddHeader("Навчальна дисципліна " + comboBox7.Text, 0, 10f, 3);
                        p.AddHeader("ЕКЗАМЕНАЦІЙНИЙ КВИТОК №" + (i + 1), 1, 10f, 2);
                        p.AddHeader("1. " + s[0], 0, 10f, 3);
                        p.AddHeader("2. " + s[1], 0, 10f, 3);
                        p.AddHeader("3. Тест", 0, 10f, 3);
                        p.AddHeader("4. Задача", 0, 10f, 3);
                        p.AddHeader("Затверджено на засіданні Кафедри комп'ютерних систем, мереж та кібербезпеки(503)", 0, 1.16f, 3);
                        p.AddHeader("протокол №" + comboBox8.Text + " від '" + comboBox9.Text + "' " + comboBox10.Text + '\x20' + comboBox11.Text + " р.", 0, 10f, 3);
                        p.AddHeader("Зав кафедри   ____________________               Екзаменатор  ____________________", 0, 10f, 3);
                        p.Newlist();
                        s.Clear();
                        ticketlistopen();
                    }
                    //закриваємо документ
                    p.Write();
                }
            }
        }
        public string _filename = "ticketlist.docx";
        /// <summary>
        /// считуємо питання
        /// </summary>
        /// <returns></returns>
        private void ticketlistopen()
        {
            //вспомогательная переменная
            bool c = false;
            //документ с питаннями
            var doc = DocX.Load(_filename);
            for (int i = 0; i < 4; i++)
            {
                //випадкове число
                int r = rnd.Next(doc.Paragraphs.Count);
                //чи є такі ж питання у одному білеті
                foreach (string a in s)
                {
                    if (doc.Paragraphs.ElementAt(r).Text == a)
                        c = true;//якщо є
                }
                if (!c)
                {
                    //якщо нема додаємо питання
                    s.Add(doc.Paragraphs.ElementAt(r).Text);
                }
                else
                {
                    //збрасуємо якщо є
                    c = false;
                    i--;
                }
            }
        }
        /// <summary>
        /// кидаємо у ворд
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (CheckComboboxesAreNotNull())
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Word w = new Word(saveFileDialog1.FileName);
                    for (int i = 0; i < NUMBERSOFTICKETS; i++)
                    {
                        //s[0 + 4 * i] (список на 100 эл для одинаковых)
                        w.AddParagraph("Національний аерокосмічний університет ім. М.Є. Жуковського 'ХАІ'", Xceed.Document.NET.Alignment.center, System.Drawing.Color.Black, true, false, 10);
                        w.AddParagraph(("Курс " + comboBox1.Text + ", ") + ("cеместр " + comboBox2.Text + ", ") + ("навчальний рік " + (comboBox3.Text) + '-' + comboBox4.Text), Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black);
                        w.AddParagraph("Спеціальність " + comboBox5.Text, Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black);
                        w.AddParagraph("Спеціальність " + comboBox6.Text, Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black);
                        w.AddParagraph("Навчальна дисципліна " + comboBox7.Text, Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.AddParagraph("ЕКЗАМЕНАЦІЙНИЙ КВИТОК №" + (i + 1), Xceed.Document.NET.Alignment.center, System.Drawing.Color.Black, true, false, 10);
                        w.AddParagraph("1. " + s[0], Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.AddParagraph("2. " + s[1], Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.AddParagraph("3. Тест", Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.AddParagraph("4. Задача", Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.AddParagraph("Затверджено на засіданні Кафедри комп'ютерних систем, мереж та кібербезпеки(503)", Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black);
                        w.AddParagraph("протокол №" + comboBox8.Text + " від '" + comboBox9.Text + "' " + comboBox10.Text + '\x20' + comboBox11.Text + " р.", Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.AddParagraph("Зав кафедри   ____________________               Екзаменатор  ____________________", Xceed.Document.NET.Alignment.left, System.Drawing.Color.Black, false, false, 10);
                        w.NewStr();
                        s.Clear();
                        ticketlistopen();
                    }
                    //зберегти
                    w.Save();
                }
            }
        }
        /// <summary>
        /// вибір питань
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                _filename = openFileDialog1.FileName;
                s.Clear();
                ticketlistopen();
            }
        }
        private bool CheckComboboxesAreNotNull()
        {
            foreach (Control c in panel1.Controls)
            {
                if (c.GetType() == typeof(ComboBox))
                    if (!(c.Name == comboBox5.Name || c.Name == comboBox6.Name))
                        if (c.Text == "")
                            return false;
            }
            return true;
        }
        /// <summary>
        /// маска для вводу даних дати
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox9_TextChanged(object sender, EventArgs e)
        {
            if (((ComboBox)sender).Text != "")
            {
                if (!Char.IsDigit(((ComboBox)sender).Text[((ComboBox)sender).Text.Length - 1]))
                {
                    if (((ComboBox)sender).Text.Length == 1)
                    {
                        ((ComboBox)sender).Text = "";
                    }
                    else
                        ((ComboBox)sender).Text = ((ComboBox)sender).Text.Remove(((ComboBox)sender).Text.Length - 1);
                }
            }
        }
    }
}
