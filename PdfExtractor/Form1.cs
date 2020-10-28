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
        string separator = ";";
        string outputfile;
        string pagesstring;
        int[] pages;
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
            if(dialog.ShowDialog()== DialogResult.OK)
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
        private bool AllIntegers(string page, char sep1, char sep2)
        {
            for (int i = 0; i < page.Length; i++)
            {
                char c = page[i];
                if(c!=sep1 && c != sep2)
                {
                    if (!Char.IsNumber(c))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private int[] GetAllNumbersBetween(int n, int n1)
        {
            List<int> r = new List<int>(); 
            for(int i=n; i<=n1; i++)
            {
                r.Add(i);
            }
            int[] ret = new int[r.Count];
            for(int i=0; i < ret.Length; i++)
            {
                ret[i] = r[i];
            }
            return ret;
        }


        private int[] GetPagesFromString(string page, char sep1, char sep2)
        {
            List<int> r = new List<int>();
            string s = "";
            for (int i = 0; i < page.Length; i++)
            {
                char c = page[i];
                if(c!=sep1 && c!= sep2)
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
                        int num1= int.Parse(s);
                        int num2;
                        int j;
                        for (j=i+1; j < page.Length; j++)
                        {
                            char cc = page[j];
                            if(cc!=sep1 && cc!= sep2)
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
                        if (j < page.Length-1)
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
                    //s = "";
                }
            }
            try
            {
                r.Add(int.Parse(s));
            }
            catch(Exception ex) { }
            int[] ret = new int[r.Count];
            for(int i=0; i < ret.Length; i++)
            {
                ret[i] = r[i];
            }
            return ret;
        }

        private string RemoveRepetedElements(string s)
        {
            string ss="";
            List<int> a = new List<int>();
            a.Clear();
            string s1 = "";
            foreach(char sss in s)
            {
                if (sss != char.Parse(separator))
                {
                    s1 += sss;
                }
                else
                {
                    a.Add(int.Parse(s1));
                    s1 = "";
                }
            }
            var dict = new Dictionary<int, int>();
            foreach(int i in a)
            {
                if (!dict.ContainsKey(i))
                {
                    dict.Add(i, 0);
                }
                dict[i]++;
            }

            foreach(var k in dict)
            {
                if (k.Value > 1)
                {
                    int v = k.Value;
                    for (int j = 0; j < v-1; j++)
                    {
                        a.RemoveAt(a.IndexOf(k.Key));
                    }
                }
            }
            foreach(int aa in a)
            {
                ss += aa.ToString() + separator;
            }
            return ss;
        }

        private void Extract_Click(object sender, EventArgs e)
        {
            outputfile = textBox2.Text;
            inputfile = textBox1.Text;
            pagesstring = textBox3.Text;
            if(outputfile.Trim()!="" && inputfile.Trim() != "" && pagesstring.Trim()!="")
            {
                if (File.Exists(inputfile))
                {
                    if (AllIntegers(pagesstring, ',', '-'))
                    {
                        pages = GetPagesFromString(pagesstring, ',', '-');
                        string ss = "";
                        foreach (int i in pages)
                        {
                            ss += i.ToString() + separator;
                        }
                        ss = RemoveRepetedElements(ss);
                        MessageBox.Show(ss, "Pages", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        PdfInfo pdfextractor = new PdfInfo(inputfile);
                        int t= pdfextractor.ExtractInfo();
                        //MessageBox.Show(t.ToString(), "Page Number", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Please put only numbers and ',' or '-'", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    }
}
