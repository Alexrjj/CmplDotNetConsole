using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace CmplConsole
{
    
    public partial class Form1 : Form
    {
        public static CheckBox Headless;
        private BackgroundWorker segPlano;

        public Form1()
        {
            InitializeComponent();

            segPlano = new BackgroundWorker();
            segPlano.DoWork += new DoWorkEventHandler(sg_DoWork);
            segPlano.RunWorkerCompleted += new RunWorkerCompletedEventHandler(sg_RunWorkerCompleted);
            button1.Click += new EventHandler(Button1_Click);
        }
        public static void Main ()
        {
            Form1 formulario = new Form1();
            Headless = formulario.checkBox1;
            Application.Run(formulario);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (!segPlano.IsBusy)
            {
                segPlano.RunWorkerAsync();
                button1.Enabled = false;
                checkBox1.Enabled = false;
            }
        }
        private void sg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                throw new Exception("Impossível continuar ", e.Error);
            }
            button1.Enabled = true;
            checkBox1.Enabled = true;
        }

        private void sg_DoWork(object sender, DoWorkEventArgs e)
        {
            ProgramaSob.Execucao();
        }
    }
}
