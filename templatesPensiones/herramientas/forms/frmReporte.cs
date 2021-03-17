using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using templatesPensiones.reportes;

namespace SISPE_MIGRACION.codigo.herramientas.forms
{
    public partial class frmReporte : Form
    {
        private bool cargando;
        private bool imprimir;
        private bool pdf;
        private string mensaje;
        private string nombrePdf;
        private tablas tablas;
        private string subcarpeta;
        internal byte[] bytes;

        public frmReporte(string nombreReporte, string subcarpeta, string tablaDataSet)
        {
            InitializeComponent();
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.p_quirogBindingSource = new System.Windows.Forms.BindingSource();
            this.tablas = new templatesPensiones.reportes.tablas();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.reportViewer1.Load += new System.EventHandler(this.reportViewer1_Load);
            this.reportViewer1.RenderingComplete += new RenderingCompleteEventHandler(camp);


            ((System.ComponentModel.ISupportInitialize)(this.p_quirogBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tablas)).BeginInit();
            this.SuspendLayout();
            this.p_quirogBindingSource.DataMember = tablaDataSet;
            this.p_quirogBindingSource.DataSource = this.tablas;
            this.tablas.DataSetName = "tablas";
            this.tablas.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DataSet1";
            reportDataSource1.Value = this.p_quirogBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = string.Format("templatesPensiones.reportes.{0}.rdlc", nombreReporte);
            this.reportViewer1.Location = new System.Drawing.Point(0, 0);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(768, 499);
            this.reportViewer1.TabIndex = 0;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 499);
            this.Controls.Add(this.reportViewer1);
            this.Name = "frmReporte";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.frmReporte_Load);

            ((System.ComponentModel.ISupportInitialize)(this.p_quirogBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tablas)).EndInit();
            this.ResumeLayout(false);
            WindowState = FormWindowState.Maximized;
            this.Text = "Reporte";
            this.ShowInTaskbar = false;


            this.subcarpeta = subcarpeta;
        }

        private void camp(object sender, RenderingCompleteEventArgs e)
        {
            string[] streamIds;
            Warning[] warnings;
            string mimeType = string.Empty;
            string encoding = string.Empty;
            string extension = string.Empty;

            if (imprimir)
            {
                this.Hide();
                if (pdf)
                {
                    bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType,
                        out encoding, out extension, out streamIds, out warnings);
                    //Se crear o se ocupa una carpeta temporal en windows sobre los pdfs

                    this.Dispose();
                }
                else
                {
                    reportViewer1.PrintDialog();
                }
                this.Close();
            }
        }

        private void reportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void frmReporte_Load(object sender, EventArgs e)
        {





        }


        internal void cargarDatos(string tablaNombre, object[] objeto, string mensaje, bool imprimir = false, object[] parametros = null)
        {



            DataTable tabla = tablas.Tables[tablaNombre];

            this.mensaje = mensaje;

            if (tabla != null)
            {

                foreach (var item in objeto)
                {
                    tabla.Rows.Add((object[])item);
                }
            }

            if (parametros != null)
            {
                object[][] elemento = (object[][])parametros;
                ReportParameter[] auxParametros = new ReportParameter[elemento[0].Length];
                for (int x = 0; x < elemento[0].Length; x++)
                {
                    ReportParameter p1 = new ReportParameter((string)elemento[0][x], (string)elemento[1][x]);
                    auxParametros[x] = p1;
                }
                reportViewer1.LocalReport.SetParameters(auxParametros);
            }
            this.reportViewer1.RefreshReport();
            this.imprimir = imprimir;
            string mimeType = string.Empty;
            string encoding = string.Empty;
            string[] streamIds;
            Warning[] warnings;
            string extension = string.Empty;
            bytes = reportViewer1.LocalReport.Render("PDF", null, out mimeType,
                       out encoding, out extension, out streamIds, out warnings);
        }
        public void setParametrosExtra(bool pdf, string nombre)
        {
            this.pdf = pdf;
            this.nombrePdf = nombre;


        }

        private void frmReporte_Load_1(object sender, EventArgs e)
        {

        }


        private void frmReporte_Load_2(object sender, EventArgs e)
        {

        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // frmReporte
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "frmReporte";
            this.Load += new System.EventHandler(this.frmReporte_Load_3);
            this.ResumeLayout(false);

        }

        private void frmReporte_Load_3(object sender, EventArgs e)
        {

        }
    }
}