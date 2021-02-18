using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PdfExtractor
{

    public partial class Form1 : Form
    {
        string inputfile;
        string outputfile;
        string pagesstring;
        List<int> pages = new List<int>();
        public Form1()
        {
            InitializeComponent();
        }

        private void InputButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = ".";
            dialog.Multiselect = false;
            dialog.Filter = "pdf files (*.pdf)|*.pdf";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
                inputfile = textBox1.Text;
            }
        }

        private void OutputButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "pdf files (*.pdf)|*.pdf";
            dialog.InitialDirectory = ".";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = dialog.FileName;
                outputfile = textBox2.Text;
            }
        }
        private bool AllIntegers(string page, char sep1, char sep2, ref bool IsZero)
        {
            IsZero = false;
            for (int i = 0; i < page.Length; i++)
            {
                char c = page[i];
                if (c != sep1 && c != sep2)
                {
                    if (!Char.IsNumber(c))
                    {
                        return false;
                    }
                    else
                    {
                        if (int.Parse(c.ToString()) < 1)
                        {
                            IsZero = true;
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        private int[] GetAllNumbersBetween(int n, int n1)
        {
            List<int> r = new List<int>();
            for (int i = n; i <= n1; i++)
            {
                r.Add(i);
            }
            int[] ret = new int[r.Count];
            for (int i = 0; i < ret.Length; i++)
            {
                ret[i] = r[i];
            }
            return ret;
        }


        private List<int> GetPagesFromString(string page, char sep1, char sep2)
        {
            List<int> r = new List<int>();
            string s = "";
            for (int i = 0; i < page.Length; i++)
            {
                char c = page[i];
                if (c != sep1 && c != sep2)
                {
                    s += c;
                }
                else
                {
                    if (c == sep1)//,
                    {
                        r.Add(int.Parse(s));
                        s = "";
                    }
                    else if (c == sep2)//-
                    {
                        string ss = "";
                        int num1 = int.Parse(s);
                        int num2;
                        int j;
                        for (j = i + 1; j < page.Length; j++)
                        {
                            char cc = page[j];
                            if (cc != sep1 && cc != sep2)
                            {
                                ss += cc;
                            }
                            else
                            {
                                break;
                            }

                        }
                        num2 = int.Parse(ss);
                        int[] nums = GetAllNumbersBetween(num1, num2);
                        r.AddRange(nums);
                        if (j < page.Length - 1)
                        {
                            s = "";
                            i = j;
                        }
                        else
                        {
                            s = "";
                            break;
                        }


                    }
                }
            }
            try
            {
                r.Add(int.Parse(s));
            }
            catch (Exception ex) { }
            return r;
        }

        private bool MoreThanOneElem(string s, char sep1, char sep2)
        {
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];

                if (c == sep1 || c == sep2)
                {
                    if (i <= s.Length - 1)
                    {
                        if (s[i + 1] == sep1 || s[i + 1] == sep2)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private bool EndOfFileIs(string s, string e)
        {
            int j = 0;
            for (int i = s.Length - e.Length; i < s.Length; i++, j++)
            {
                if (s[i] != e[j])
                {
                    return false;
                }
            }
            return true;
        }

        private void Extract_Click(object sender, EventArgs e)
        {
            outputfile = textBox2.Text;
            inputfile = textBox1.Text;
            pagesstring = textBox3.Text;
            if (outputfile.Trim() != "" && inputfile.Trim() != "" && pagesstring.Trim() != "")
            {
                if (!EndOfFileIs(outputfile, ".pdf"))
                {
                    outputfile += ".pdf";
                }

                if (!EndOfFileIs(inputfile, ".pdf"))
                {
                    inputfile += ".pdf";
                }

                if (File.Exists(inputfile))
                {
                    bool IsZero = false;
                    if (AllIntegers(pagesstring, ',', '-', ref IsZero))
                    {
                        if (!MoreThanOneElem(pagesstring, ',', '-'))
                        {
                            pages.Clear();
                            pages = GetPagesFromString(pagesstring, ',', '-');
                            PdfInfo pdfextractor = new PdfInfo(inputfile);
                            int t = pdfextractor.ExtractInfo();

                            bool allOk = true;
                            foreach (int i in pages)
                            {
                                if (i > t)
                                {
                                    allOk = false;
                                    MessageBox.Show("Pdf does not have more than " + i + " pages !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                            }
                            if (allOk)
                            {
                                //Extract pages
                                ExtractPages(inputfile, outputfile, pages);
                                MessageBox.Show("Pages extracted !", "Congratulations!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            MessageBox.Show(", or - are more than one times beside !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    }
                    else
                    {
                        if (IsZero)
                        {
                            MessageBox.Show("Please put numbers bigger or equal to 1 !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            MessageBox.Show("Please put only numbers and ',' or '-'", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Input file: " + inputfile + " does not exist", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please fill the pages and the input and output file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void HelpButton_Click(object sender, EventArgs e)
        {
            HelpDialog dialog = new HelpDialog();
            dialog.ShowDialog();
        }

        private void ExtractPages(string inputfilename, string outputfilename, List<int> pages)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the 
                // contents of the source Pdf file:
                reader = new PdfReader(inputfilename);


                sourceDocument = new Document();

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument,
                    new System.IO.FileStream(outputfilename, System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the array and add the page copies to the output file:
                foreach (int pageNumber in pages)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, pageNumber);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
