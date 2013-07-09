using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Table_Project
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            //Data.DataLink getmedata4 = new Data.DataLink();
            //getmedata4.julie();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox1.Text = "";
                    textBox1.Text = openFileDialog1.FileName.ToString();
                    //if ((myStream = openFileDialog1.OpenFile()) != null)
                    //{
                    //    using (myStream)
                    //    {
                    //        // Insert code to read the stream here.
                    //    }
                    //}
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }



        }

        private void button1_Click(object sender, EventArgs e)
        {
             Data.DataLink getmedata4 = new Data.DataLink();
             getmedata4.exelinfo(textBox1.Text, textBox2.Text);
             MessageBox.Show("done");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.doupdateme();
            MessageBox.Show("done");
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.testing(textBox3.Text.Trim());
            MessageBox.Show("done");
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.getsecond();
            MessageBox.Show("done");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.matchpieceno12();
            MessageBox.Show("done");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.stockupdate();
            MessageBox.Show("done");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Data.DataLink getmedata4 = new Data.DataLink();
            getmedata4.stockpiece();
            MessageBox.Show("done");
            
        }
    }
}
