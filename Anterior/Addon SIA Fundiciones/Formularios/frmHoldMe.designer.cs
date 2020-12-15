namespace Addon_SIA
{
    partial class frmHoldMe
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.TimerAccountSetup = new System.Windows.Forms.Timer(this.components);
            this.TimerCrystalReports = new System.Windows.Forms.Timer(this.components);
            this.TimerManageReports = new System.Windows.Forms.Timer(this.components);
            this.TimerManageGroups = new System.Windows.Forms.Timer(this.components);
            this.TimerManageGrpUsers = new System.Windows.Forms.Timer(this.components);
            this.TimerCrystalViewer = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // TimerCrystalViewer
            // 
            this.TimerCrystalViewer.Tick += new System.EventHandler(this.TimerCrystalViewer_Tick);
            // 
            // frmHoldMe
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(219, 107);
            this.IsMdiContainer = true;
            this.MaximizeBox = false;
            this.Name = "frmHoldMe";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Crystal Reports";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmHoldMe_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer TimerAccountSetup;
        private System.Windows.Forms.Timer TimerCrystalReports;
        private System.Windows.Forms.Timer TimerManageReports;
        private System.Windows.Forms.Timer TimerManageGroups;
        private System.Windows.Forms.Timer TimerManageGrpUsers;
        private System.Windows.Forms.Timer TimerCrystalViewer;
    }
}