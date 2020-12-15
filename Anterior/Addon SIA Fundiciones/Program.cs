using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SAPbouiCOM;
using System.Threading;

namespace Addon_SIA
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
//        [STAThread]
        
        [STAThread]
        
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            //csUtilidades csUtilidades = new csUtilidades();
            try
            {
                if (csUtilidades.Inicio())
                {
                    csVariablesGlobales.SboApp.StatusBar.SetText("Add-On SIA cargado correctamente", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                    frmHoldMe FrmHoldMe = new frmHoldMe();
                    FrmHoldMe.Visible = false;
                    System.Windows.Forms.Application.Run();
                }
                else
                {
                    csVariablesGlobales.SboApp.StatusBar.SetText("El Add-On SIA no se ha cargado correctamente", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            catch
            {
                csVariablesGlobales.SboApp.StatusBar.SetText("El Add-On SIA no se ha cargado correctamente", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}