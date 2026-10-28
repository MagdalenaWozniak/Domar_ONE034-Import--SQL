using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace one017
{
    public partial class WyborSciezek : Form
    {
        public WyborSciezek()
        {
            InitializeComponent();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox1.Text = fbd.SelectedPath;
        }
        private void TextBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox2.Text = fbd.SelectedPath;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            string sciezka1 = "";
            string sciezka2 = "";

            Cursor.Current = Cursors.WaitCursor;

            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("Nie podno lokalizacji.", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                sciezka1 = textBox1.Text.Replace("\\", "\\\\");
                sciezka2 = textBox2.Text.Replace("\\", "\\\\");
            }





            //MessageBox.Show(sciezka+" "+sciezkatemp);
            try
            {


                var komenda = "\""+sciezka1+"\" \""+sciezka2+"\"";
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.UseShellExecute = false;
                try
                {
                    process.Start();
                    process.StandardInput.WriteLine(komenda);
                    process.StandardInput.Flush();
                    process.StandardInput.Close();
                    process.Close();

                    //MessageBox.Show(process.StandardOutput.ReadToEnd());

                }
                catch (Exception expe)
                {
                    MessageBox.Show("Błąd przy egzekwowaniu wpisu do wiersza poleceń: " + expe);
                }

                return;
            }
            catch(Exception ee)
            {
                MessageBox.Show("Błąd: "+ ee.Message );
            }
        }

    }
}

