
namespace ReportePeriodo
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.rvDiario = new Microsoft.Reporting.WinForms.ReportViewer();
            this.etiEstablo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Location = new System.Drawing.Point(18, 76);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 0;
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
            // 
            // rvDiario
            // 
            this.rvDiario.LocalReport.ReportEmbeddedResource = "ReportePeriodo.ReporteDiario.rdlc";
            this.rvDiario.Location = new System.Drawing.Point(0, 0);
            this.rvDiario.Name = "rvDiario";
            this.rvDiario.ServerReport.BearerToken = null;
            this.rvDiario.Size = new System.Drawing.Size(396, 246);
            this.rvDiario.TabIndex = 0;
            // 
            // etiEstablo
            // 
            this.etiEstablo.AutoSize = true;
            this.etiEstablo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.etiEstablo.Location = new System.Drawing.Point(15, 22);
            this.etiEstablo.Name = "etiEstablo";
            this.etiEstablo.Size = new System.Drawing.Size(76, 15);
            this.etiEstablo.TabIndex = 1;
            this.etiEstablo.Text = "ESTABLO: ";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(293, 256);
            this.Controls.Add(this.etiEstablo);
            this.Controls.Add(this.monthCalendar1);
            this.Name = "Form1";
            this.Text = "Reporte Diario";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private Microsoft.Reporting.WinForms.ReportViewer rvDiario;
        private System.Windows.Forms.Label etiEstablo;
    }
}

