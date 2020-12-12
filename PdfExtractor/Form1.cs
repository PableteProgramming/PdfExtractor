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
        private bool IsRepeatedElements(string s, List<string> r)
        {
            r.Clear();
            List<string> elems= new List<string>();
            elems.Clear();

            string ss = "";

            foreach(char c in s)
            {
                if (c.ToString() != separator)
                {
                    ss += c;
                }
                else
                {
                    elems.Add(ss);
                    ss = "";
                }
            }
            elems.Add(ss);

            IEnumerable<string> duplicates = elems.GroupBy(x => x)
                                        .Where(g => g.Count() > 1)
                                        .Select(x => x.Key);

            if (duplicates.Count() > 0)
            {
                foreach(string st in duplicates)
                {
                    r.Add(st);
                }

                return true;
            }

            return false;
        }

        private bool MoreThanOneElem(string s, char sep1, char sep2)
        {
            //string pattern = $@"[{sep1}][{sep1}*]";
            //string pattern1 = $@"[{sep1}][{sep2}*]";
            //string pattern2 = $@"[{sep2}][{sep1}*]";
            //string pattern3 = $@"[{sep2}][{sep2}*]";
            for(int i=0; i<s.Length;i++)
            {
                char c = s[i];

                if(c==sep1 || c == sep2)
                {
                    if (i <= s.Length - 1)
                    {
                        if (s[i + 1] ==sep1 || s[i + 1] == sep2)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private void RunPythonProgram(string output,string input, string pages)
        {
            if (!File.Exists("extract.exe"))
            {
                MessageBox.Show("extract.exe is needed and is not found ! Please put it in the same directory as the executable. You can always find it here: https://github.com/PableteProgramming/extractpdf", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "cmd.exe";
                startInfo.CreateNoWindow = true;
                startInfo.Arguments = "/C extract.exe " + input + " " + output+" "+pages;
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                MessageBox.Show("Pages "+pages +" extracted !", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
                        //Error

                        if (!MoreThanOneElem(pagesstring, ',', '-'))
                        {
                            pages = GetPagesFromString(pagesstring, ',', '-');
                            string ss = "";
                            foreach (int i in pages)
                            {
                                ss += i.ToString() + separator;
                            }
                            List<string> repeatedElems = new List<string>();
                            repeatedElems.Clear();
                            if (IsRepeatedElements(ss, repeatedElems))
                            {
                                bool allOk = false;
                                foreach (string rs in repeatedElems)
                                {
                                    if (MessageBox.Show("Are you sure you want to have multiples times the page " + rs + " ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        allOk = true;
                                    }
                                    else
                                    {
                                        allOk = false;
                                    }
                                    if (!allOk)
                                    {
                                        break;
                                    }
                                }

                                if (allOk)
                                {

                                    PdfInfo pdfextractor = new PdfInfo(inputfile);
                                    int t = pdfextractor.ExtractInfo();

                                    allOk = true;
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
                                        //MessageBox.Show(ss, "Pages", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        RunPythonProgram(outputfile.Trim(), inputfile.Trim(), ss);
                                        ///////Run python program
                                    }
                                }

                            }
                            else
                            {
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
                                    //MessageBox.Show(ss, "Pages", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    RunPythonProgram(outputfile.Trim(), inputfile.Trim(), ss);
                                    ///////Run python program
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show(", or - are more than one times beside !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        
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
