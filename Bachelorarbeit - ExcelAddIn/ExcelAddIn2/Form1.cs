using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using static ExcelAddIn2.MyRibbon;
using Newtonsoft.Json;
using System.IO;

namespace ExcelAddIn2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            AcceptButton = button1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /*
            textBox1.Width = 250;
            textBox1.Height = 50;
            textBox1.Multiline = true;
            textBox1.BackColor = Color.Blue;
            textBox1.ForeColor = Color.White;
            textBox1.BorderStyle = BorderStyle.Fixed3D;
            */
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string var;
            var = textBox1.Text;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string userInput1 = textBox1.Text;

            Debug.WriteLine("userInput e " + userInput1);

            //setCharacteristics(userInput1);
            setCharacteristics();

            string userInput2 = textBox2.Text;

            Debug.WriteLine("userInput e " + userInput2);
            //MyRibbon.RunAsync(userInput1, userInput2).Wait();
            MyRibbon.setUrl1(userInput1);
            MyRibbon.setUrl2(userInput2);
        }

    }
}
