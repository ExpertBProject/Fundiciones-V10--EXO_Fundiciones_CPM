namespace Addon_SIA
{
    partial class frmDialog : System.Windows.Forms.Form
    {

        //Form reemplaza a Dispose para limpiar la lista de componentes.
        [System.Diagnostics.DebuggerNonUserCode()]
        protected override void Dispose(bool disposing) {
        if( disposing && components != null ){
            components.Dispose();
        }
        base.Dispose(disposing);
    }

        //Requerido por el Diseñador de Windows Forms
        private System.ComponentModel.IContainer components = null;

        //NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
        //Se puede modificar usando el Diseñador de Windows Forms.
        //No lo modifique con el editor de código.
        [System.Diagnostics.DebuggerStepThrough()]
        private void InitializeComponent()
        {
            this.PrintDialog1 = new System.Windows.Forms.PrintDialog();
            this.FolderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.OpenFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SaveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            //
            //PrintDialog1
            //
            this.PrintDialog1.UseEXDialog = true;
            //
            //OpenFileDialog1
            //
            this.OpenFileDialog1.FileName = "OpenFileDialog1";
            //
            //FrmDialog
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6.0F, 13.0F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(268, 135);
            this.Name = "FrmDialog";
            this.ResumeLayout(false);

        }
        internal System.Windows.Forms.PrintDialog PrintDialog1;
        internal System.Windows.Forms.FolderBrowserDialog FolderBrowserDialog1;
        internal System.Windows.Forms.OpenFileDialog OpenFileDialog1;
        internal System.Windows.Forms.SaveFileDialog SaveFileDialog1;
    }
}