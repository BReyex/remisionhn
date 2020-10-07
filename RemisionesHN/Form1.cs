using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RemisionesHN
{
    public partial class Form1 : Form
    {
        public static class Remision
        {
            public static string CodRemision;
        }
        public static class Servidor
        {
            public static string urlServidor;
        }
        public static class Base
        {
            public static string Basedatos;
        }
        public Form1(string corem,string arg, string bd)
        {
            Remision.CodRemision = corem;
            Servidor.urlServidor = arg;
            Base.Basedatos = bd;
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

         

            using (SqlConnection conn2 = new SqlConnection("Data Source="+Servidor.urlServidor+";Initial Catalog="+Base.Basedatos+";Persist Security Info=True;User ID=sa;Password=jda"))
            {


                string sql = @"SELECT f.FACTURA FROM emasalhn.FACTURA f WHERE  f.Tipo_Documento='R' and f.PEDIDO = (SELECT p.PEDIDO FROM emasalhn.PEDIDO p WHERE p.RowPointer='" + Remision.CodRemision.ToString() + "') ";
                using (var command = new SqlCommand(sql, conn2))
                {
                    conn2.Open();
                    try
                    {
                        var reader = command.ExecuteScalar();
                        Remision.CodRemision = reader.ToString();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("EL pedido no hace referencia a una remision");
                        this.Close();
                    }
                   
                    
                       
                    
                }
            }

                string NombreImpresora = "";

                ////Busco la impresora por defecto
                for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
                {
                    PrinterSettings a = new PrinterSettings();
                    a.PrinterName = PrinterSettings.InstalledPrinters[i].ToString();
                    if (a.IsDefaultPrinter)
                    {
                        NombreImpresora = PrinterSettings.InstalledPrinters[i].ToString();

                    }
                }

                int i2 = 1;
            while (i2 <= 3)
            {
                ReportDocument cryRpt;
                cryRpt = new ReportDocument();
                cryRpt.Load(@"C:\RemisionesHN\REMISIONHN.rpt");
                //cryRpt.DataSourceConnections[0].SetConnection("192.168.7.2", "EXACTUS", true);
                cryRpt.SetParameterValue("sTipoCopia", i2.ToString());
                cryRpt.SetParameterValue("Factura", Remision.CodRemision.ToString());

                cryRpt.SetDatabaseLogon("sa", "jda");
                cryRpt.PrintOptions.PrinterName = NombreImpresora; //Asigno la impresora
                cryRpt.PrintToPrinter(1, false, 0, 0);//Imprimo 2 copias
                i2++;
            }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


            this.Close();
            // crystalReportViewer1.ReportSource = cryRpt;
            // crystalReportViewer1.Refresh();



            //ReportDocument cryRpt = new ReportDocument();
            //TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
            //TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
            //ConnectionInfo crConnectionInfo = new ConnectionInfo();
            //Tables CrTables;

            //cryRpt.Load(@"C:\SoftlandERP\Reportes\REMISIONHN.rpt");

            //crConnectionInfo.ServerName = "192.168.7.2";
            //crConnectionInfo.DatabaseName = "EXACTUS";
            //crConnectionInfo.UserID = "sa";
            //crConnectionInfo.Password = "jda";

            //CrTables = cryRpt.Database.Tables;
            //foreach (CrystalDecisions.CrystalReports.Engine.Table CrTable in CrTables)
            //{
            //    crtableLogoninfo = CrTable.LogOnInfo;
            //    crtableLogoninfo.ConnectionInfo = crConnectionInfo;
            //    CrTable.ApplyLogOnInfo(crtableLogoninfo);
            //}
            //cryRpt.SetParameterValue("sTipoCopia", "1");
            //cryRpt.SetParameterValue("Factura", "REM15449");
            //crystalReportViewer1.ReportSource = cryRpt;
            //crystalReportViewer1.Refresh();
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {

        }
    }
}
