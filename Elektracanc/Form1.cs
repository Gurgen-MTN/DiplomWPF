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
using System.Data.Sql;
using System.Data.SqlClient;
using Microsoft.Win32;
using System.IO;


namespace Elektracanc
{
    public partial class Form1 : Form
    {

        
        public Form1()
        {
            InitializeComponent();
        }
        private async void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void sozdanieToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Moduli_Sozdanie f = new Moduli_Sozdanie();
            f.Show();

        }

        private void korrektirovkaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Moduli_izmenenie f = new Moduli_izmenenie();
            f.Show();
        }

        private void udalenieToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Moduli.Moduli_udalenie f = new Moduli.Moduli_udalenie();
            f.Show();
        }




        private void sozdanieToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Schetchiki.Schetchiki_sozdanie f = new Schetchiki.Schetchiki_sozdanie();
            f.Show();
        }

        private void udalenieToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Schetchiki.Schetchiki_udalenie f = new Schetchiki.Schetchiki_udalenie();
            f.Show();
        }

        private void izmenenieToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Schetchiki.Schetchiki_izmenenie f = new Schetchiki.Schetchiki_izmenenie();
            f.Show();
        }

        private void moduleiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PoiskIOtechtei.PoiskModuleysRazryajennimiBatareykami f = new PoiskIOtechtei.PoiskModuleysRazryajennimiBatareykami();
            f.Show();
        }

        private void moduleySNesankcionirovannimDostupomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PoiskIOtechtei.ModuleySNesankcionirovannimDostupom f = new PoiskIOtechtei.ModuleySNesankcionirovannimDostupom();
            f.Show();
        }



        private void informaciiOSchetchikeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PoiskIOtechtei.InformaciiOSchetchike f = new PoiskIOtechtei.InformaciiOSchetchike();
            f.Show();
        }

        private void vvodFaylovToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VvodDannix.VvodFaylov f = new VvodDannix.VvodFaylov();
            f.Show();
        }


        private void oProgrammeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Pomosh.oProgramme f = new Pomosh.oProgramme();
            f.Show();
        }

        private void schetchikovSIstekshimSrokomProverkiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PoiskIOtechtei.SchetchikovSIstekshimSrokomProverki f = new PoiskIOtechtei.SchetchikovSIstekshimSrokomProverki();
            f.Show();
        }
    }
}
