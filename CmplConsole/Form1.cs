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
        private BackgroundWorker segPlano = new BackgroundWorker();
        private BackgroundWorker geraPedidoSap = new BackgroundWorker();
        private BackgroundWorker extraiStatus = new BackgroundWorker();


        public Form1()
        {
            InitializeComponent();

            segPlano.DoWork += new DoWorkEventHandler(sg_VistoriaSob);
            geraPedidoSap.DoWork += new DoWorkEventHandler(sg_GeraPedidoSap);
            extraiStatus.DoWork += new DoWorkEventHandler(sg_ExtraiStatus);
            segPlano.RunWorkerCompleted += new RunWorkerCompletedEventHandler(sg_RunWorkerCompleted);
            geraPedidoSap.RunWorkerCompleted += new RunWorkerCompletedEventHandler(sg_RunWorkerCompleted);
            extraiStatus.RunWorkerCompleted += new RunWorkerCompletedEventHandler(sg_RunWorkerCompleted);
            
            button1.Click += new EventHandler(Button1_Click);
            button2.Click += new EventHandler(Button2_Click);
            Button3.Click += new EventHandler(Button3_Click);

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

        private void sg_VistoriaSob(object sender, DoWorkEventArgs e)
        {
            ProgramaSob.Vistoria();
        }

        private void sg_GeraPedidoSap(object sender, DoWorkEventArgs e)
        {
            ProgramaSob.GeraPedSAP();
        }

        private void sg_ExtraiStatus(object sender, DoWorkEventArgs e)
        {
            Extrair.SobStatus();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (!geraPedidoSap.IsBusy)
            {
                geraPedidoSap.RunWorkerAsync();
                button1.Enabled = false;
                checkBox1.Enabled = false;
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (!extraiStatus.IsBusy)
            {
                extraiStatus.RunWorkerAsync();
                button1.Enabled = false;
                checkBox1.Enabled = false;
                button2.Enabled = false;
            }
        }
    }
}
