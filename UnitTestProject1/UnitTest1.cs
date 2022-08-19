using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace UnitTestProject1
{
    using tickets;
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Word w;
            var exception = Assert.ThrowsException<Exception>(() => w = new Word(""));
        }
        [TestMethod]
        public void TestMethod2()
        {
            PdfWrite w;
            var exception = Assert.ThrowsException<Exception>(() => w = new PdfWrite(""));
        }
        [TestMethod]
        public void TestMethod3()
        {
            Word w = new Word("out.docx");
            w.Save();
        }
        [TestMethod]
        public void TestMethod4()
        {
            PdfWrite w = new PdfWrite("out.pdf");
            w.AddHeader("Національний аерокосмічний університет ім. М.Є. Жуковського 'ХАІ'", 1, 10f, 2);
            w.Write();
        }
    }
}
