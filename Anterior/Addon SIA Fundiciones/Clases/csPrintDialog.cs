using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Addon_SIA
{
    class csPrintDialog
    {
        public string StrNomImp;
        public short NumCop;

        public System.Windows.Forms.DialogResult Impresora(short NumeroCopias)
        {
            frmDialog FrmDialog = new frmDialog();
            FrmDialog.PrintDialog1.PrinterSettings.Copies = NumeroCopias;
            System.Windows.Forms.DialogResult dialogResult;
            dialogResult = FrmDialog.PrintDialog1.ShowDialog();
            FrmDialog.TopMost = true;
            FrmDialog.BringToFront();
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                StrNomImp = FrmDialog.PrintDialog1.PrinterSettings.PrinterName;
                NumCop = FrmDialog.PrintDialog1.PrinterSettings.Copies;
            }
            return dialogResult;
        }

    }
}
