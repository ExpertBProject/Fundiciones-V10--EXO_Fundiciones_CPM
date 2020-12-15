using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
using System.ComponentModel;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Addon_SIA
{
    class csFuncionalidadSAP
    {
        SAPbouiCOM.Form oForm = null;
        SAPbouiCOM.EditText oEditText = null;
        SAPbouiCOM.StaticText oStaticText = null;
        SAPbouiCOM.Matrix oMatrix = null;
        SAPbouiCOM.Item oItem = null;
        SAPbouiCOM.Button oButton = null;
        SAPbouiCOM.Column oColumn = null;
        SAPbouiCOM.Folder oFolder = null;
        SAPbouiCOM.CheckBox oCheckBox = null;
        SAPbouiCOM.ComboBox oComboBox = null;

        int FilaSeleccionada;
        
        public void SetFilters()
        {
            csVariablesGlobales.SboApp.StatusBar.SetText("Estableciendo Filters SAP", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);
            try
            {
                //csUtilidades csUtilidades = new csUtilidades();
                csVariablesGlobales.oFilters = new SAPbouiCOM.EventFilters();
                #region Formularios SAP
                //csVariablesGlobales.oFilter.AddEx("41"); //Selección de Lote en Formularios SAP
                //csVariablesGlobales.oFilter.AddEx("133"); //Factura de Ventas
                //csVariablesGlobales.oFilter.AddEx("134"); //Interlocutores Comerciales
                //csVariablesGlobales.oFilter.AddEx("139"); //Pedido de Ventas
                //csVariablesGlobales.oFilter.AddEx("140"); //Entregas
                //csVariablesGlobales.oFilter.AddEx("149"); //Oferta
                //csVariablesGlobales.oFilter.AddEx("180"); //Devoluciones
                //csVariablesGlobales.oFilter.AddEx("181"); //Abono Proveedores
                //csVariablesGlobales.oFilter.AddEx("182"); //Devolucion de Compras
                //csVariablesGlobales.oFilter.AddEx("179"); //Abono Deudores
                //csVariablesGlobales.oFilter.AddEx("142"); //Pedido de Compras
                //csVariablesGlobales.oFilter.AddEx("143"); //Entradas
                //csVariablesGlobales.oFilter.AddEx("141"); //A/P Invoice
                //csVariablesGlobales.oFilter.AddEx("181"); //A/P Credit Memo 
                //csVariablesGlobales.oFilter.AddEx("392"); //Asientos
                //csVariablesGlobales.oFilter.AddEx("606"); //Depósito
                //csVariablesGlobales.oFilter.AddEx("720"); //Salida Mercancías
                //csVariablesGlobales.oFilter.AddEx("721"); //Entrada Mercancías
                //csVariablesGlobales.oFilter.AddEx("3002"); //Documentos Preliminares
                //csVariablesGlobales.oFilter.AddEx("60052"); //Operaciones de Efectos
                //csVariablesGlobales.oFilter.AddEx("60090"); //A/R Invoice - Factura Clientes + Pago
                //csVariablesGlobales.oFilter.AddEx("60091"); //A/R Reserve Invoice - Factura Anticipo Clientes
                //csVariablesGlobales.oFilter.AddEx("60092"); //A/P Reserve Invoice
                //csVariablesGlobales.oFilter.AddEx("65300"); //A/R Down Payment Invoice, - Factura Anticipo Deudores
                //csVariablesGlobales.oFilter.AddEx("65301"); //A/P Down Payment Invoice, 
                //csVariablesGlobales.oFilter.AddEx("65308"); //A/R Down Payment Request 
                //csVariablesGlobales.oFilter.AddEx("65309"); //A/P Down Payment Request 
                //csVariablesGlobales.oFilter.AddEx("65211"); //Orden de Fabricación
                //csVariablesGlobales.oFilter.AddEx("10003"); //Seleccion Articulos
                //csVariablesGlobales.oFilter.AddEx("65010"); //Informe 347
                #endregion
                #region Filters
                #region Item Pressed
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
                #endregion
                #region Menu Click
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_MENU_CLICK);
                csVariablesGlobales.oFilter.AddEx("ALL_FORMS");
                #endregion
                #region Click
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
                if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSAP", "U_FormSIA='@" + csVariablesGlobales.Prefijo + "_FORMS'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSAP", "U_FormSIA='@" + csVariablesGlobales.Prefijo + "_FORMS'")); //
                }
                if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSAP", "U_FormSIA='@" + csVariablesGlobales.Prefijo + "_REPORT'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSAP", "U_FormSIA='@" + csVariablesGlobales.Prefijo + "_REPORT'")); //
                }
                if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSAP", "U_FormSIA='@" + csVariablesGlobales.Prefijo + "_PARAM'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSAP", "U_FormSIA='@" + csVariablesGlobales.Prefijo + "_PARAM'")); //
                }
                csVariablesGlobales.oFilter.AddEx("SIASL00001"); //FrmImportacionesSIA
                csVariablesGlobales.oFilter.AddEx("SIASL00002"); //FrmImpresionListados
                csVariablesGlobales.oFilter.AddEx("SIASL00003"); //FrmAsientoGastos
                csVariablesGlobales.oFilter.AddEx("606"); //Depósito
                csVariablesGlobales.oFilter.AddEx("65211"); //Orden Producción
                csVariablesGlobales.oFilter.AddEx("SIASL10001"); //FrmEntSalStock  
                csVariablesGlobales.oFilter.AddEx("SIASL10002"); //FrmProcesos
                csVariablesGlobales.oFilter.AddEx("SIASL10005"); //FrmModelo347
                //csVariablesGlobales.oFilter.AddEx("2001060004"); //Seleccionar Informe 133
                //csVariablesGlobales.oFilter.AddEx("2001060005"); //Seleccionar Informe 140
                //csVariablesGlobales.oFilter.AddEx("2001060006"); //Seleccionar Informe 149
                csVariablesGlobales.oFilter.AddEx("SIASL10003"); //FrmGeneraMemoria
                //csVariablesGlobales.oFilter.AddEx("2001060011"); //Seleccionar Informe 142
                csVariablesGlobales.oFilter.AddEx("SIASL10006"); //FrmModelo349
                //csVariablesGlobales.oFilter.AddEx("2001060014"); //Seleccionar Informe 139
                csVariablesGlobales.oFilter.AddEx("392"); //Asientos
                //csVariablesGlobales.oFilter.AddEx("2001060015"); //FrmAsientoGastos
                csVariablesGlobales.oFilter.AddEx("3002");
                //csVariablesGlobales.oFilter.AddEx("2001060017"); //FrmImportacionesSIA
                //csVariablesGlobales.oFilter.AddEx("2001060018"); //Seleccionar Informe 3002
                //csVariablesGlobales.oFilter.AddEx("2001060019"); //Seleccionar Informe 60052
                //csVariablesGlobales.oFilter.AddEx("2001060020"); //Seleccionar Informe 180
                //csVariablesGlobales.oFilter.AddEx("2001060021"); //Seleccionar Informe 179
                csVariablesGlobales.oFilter.AddEx("65011"); //Criterios de Selccion 347
                #endregion
                #region Double Click
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
                csVariablesGlobales.oFilter.AddEx("SIASL00004"); //FrmImpReports
                #endregion
                #region Lost Focus
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
                csVariablesGlobales.oFilter.AddEx("SIASL10001"); //FrmEntSalStock
                csVariablesGlobales.oFilter.AddEx("SIASL10002"); //FrmProcesos
                if (csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_PLANO'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_PLANO'")); //
                }
                if (csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_UBIC'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_UBIC'")); //
                }
                if (csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_REPORT'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_REPORT'")); //
                }
                if (csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_PEDC'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_PEDC'")); //
                }
                csVariablesGlobales.oFilter.AddEx("41"); //
                #endregion
                #region Got Focus
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
                csVariablesGlobales.oFilter.AddEx("SIASL10002"); //FrmProcesos
                csVariablesGlobales.oFilter.AddEx("41"); //
                #endregion
                #region Choose From List
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
                csVariablesGlobales.oFilter.AddEx("SIASL00003"); //FrmAsientoGastos

                csVariablesGlobales.oFilter.AddEx("SIASL10001"); //FrmEntSalStock  
                csVariablesGlobales.oFilter.AddEx("SIASL10002"); //FrmProcesos  
                csVariablesGlobales.oFilter.AddEx("SIASL10005"); //FrmModelo347
                csVariablesGlobales.oFilter.AddEx("SIASL10006"); //FrmModelo349
                //csVariablesGlobales.oFilter.AddEx("2001060015"); //FrmAsientoGastos
                #endregion
                #region Key Down
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
                csVariablesGlobales.oFilter.AddEx("SIASL10002"); //FrmProcesos
                #endregion
                #region Right Click
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_RIGHT_CLICK);
                if (csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_UBIC'") != "")
                {
                    csVariablesGlobales.oFilter.AddEx(csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSAP", "U_FormSIA='@SIA_UBIC'")); //
                }
                #endregion
                #region Form Load
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_LOAD);
                csVariablesGlobales.oFilter.AddEx("65211"); //Orden Producción
                csVariablesGlobales.oFilter.AddEx("134"); //Interlocutores Comerciales
                csVariablesGlobales.oFilter.AddEx("606"); //Depósito
                csVariablesGlobales.oFilter.AddEx("392"); //Asientos
                csVariablesGlobales.oFilter.AddEx("3002"); //Documentos Preliminares
                csVariablesGlobales.oFilter.AddEx("10003"); //Seleccion Articulos
                csVariablesGlobales.oFilter.AddEx("65010"); //Informe 347
                csVariablesGlobales.oFilter.AddEx("60051"); //Gestión de Efectos
                csVariablesGlobales.oFilter.AddEx("721"); //Entrada Mercancías
                csVariablesGlobales.oFilter.AddEx("720"); //Salida Mercancías
                #endregion
                #region Form Unload
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_UNLOAD);
                #endregion
                #region Form Data Add
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_DATA_ADD);
                csVariablesGlobales.oFilter.AddEx("133"); //Factura de Ventas
                csVariablesGlobales.oFilter.AddEx("179"); //Abono Deudores
                csVariablesGlobales.oFilter.AddEx("60090"); //A/R Invoice - Factura Clientes + Pago
                csVariablesGlobales.oFilter.AddEx("60091"); //A/R Reserve Invoice - Factura Anticipo Clientes
                csVariablesGlobales.oFilter.AddEx("65300"); //A/R Down Payment Invoice, 65308 A/R Down Payment - Factura Anticipo Deudores
                #endregion
                #region Form Data Load
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_DATA_LOAD);
                #endregion
                #region Form Activate
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_ACTIVATE);
                csVariablesGlobales.oFilter.AddEx("392"); //Asientos
                #endregion
                #region Form Deactivate
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_DEACTIVATE);
                #endregion
                #region Form Resize
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_FORM_RESIZE);
                #endregion
                #region Form Close
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
                #endregion
                #region Validate
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
                #endregion
                #region Form Resize
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
                #endregion
                #region Combo Select
                csVariablesGlobales.oFilter = csVariablesGlobales.oFilters.Add(BoEventTypes.et_COMBO_SELECT);
                csVariablesGlobales.oFilter.AddEx("133");
                csVariablesGlobales.oFilter.AddEx("179");
                #endregion
                #endregion
                csVariablesGlobales.SboApp.SetFilter(csVariablesGlobales.oFilters);
            }
            catch
            {
                csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetLastErrorDescription(), 1, "", "", "");
            }
        }

        public void EventFilter()
        {
            //csUtilidades csUtilidades = new csUtilidades();
            csUtilidades.CargarMenu();
            SetFilters();

            csVariablesGlobales.SboApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBOApp_AppEvent);
            csVariablesGlobales.SboApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBOApp_MenuEvent);
            csVariablesGlobales.SboApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBOApp_ItemEvent);
            csVariablesGlobales.SboApp.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBOApp_RightClickEvent);
            csVariablesGlobales.SboApp.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBOApp_FormDataEvent);
        }

        private void SBOApp_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm;
            oForm = csVariablesGlobales.SboApp.Forms.ActiveForm;
            string Formulario = csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSIA", "U_FormSAP='" + oForm.Type.ToString() + "'");
            switch (Formulario)
            {
                case "@SIA_UBIC":
                    if (eventInfo.BeforeAction == true && eventInfo.ActionSuccess == false)
                    {
                        FilaSeleccionada = eventInfo.Row;
                        //BubbleEvent = false;
                    }
                    break;
            }
        }

        private void SBOApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //csUtilidades csUtilidades = new csUtilidades();
            csImpresiones Impresiones = new csImpresiones();
            //if (csVariablesGlobales.DirectorioActual != Environment.CurrentDirectory)
            //{
            //    Environment.CurrentDirectory = csVariablesGlobales.DirectorioActual;
            //}
            string Formulario = "";
            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case "SubMnuImpLis":
                            csFrmImpresionListados FrmImpresionListados = new csFrmImpresionListados();
                            //FrmImpresionListados.CargarFormulario();
                            Formulario = "FrmImpresionListados";
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath +
                                                   @"\Formularios\FrmImpresionListados.xml",
                                                   csVariablesGlobales.SboApp);
                            FrmImpresionListados.CargaMatrix();
                            //csFrmImpReports FrmImpReports = new csFrmImpReports();
                            //FrmImpReports.CargarFormulario();
                            break;
                        case "SubMnuImpSIA":
                            csFrmImportacionesSIA FrmImportacionesSIA = new csFrmImportacionesSIA();
                            //FrmImportacionesSIA.CargarFormulario();
                            Formulario = "FrmImportacionesSIA";
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath +
                                                   @"\Formularios\FrmImportacionesSIA.xml", 
                                                   csVariablesGlobales.SboApp);
                            break;
                        case "519": //Imprimir por Pantalla
                        case "520": //Imprimir por Impresora
                        case "7176": //Exportar a Pdf:
                        //case "7169": //Exportar a Excel
                        case "7170": //Exportar a Word
                            Impresiones.Impresion(ref csVariablesGlobales.SboApp, ref pVal, out BubbleEvent);
                            if (BubbleEvent == false)
                            {
                                return;
                            }
                            break;
                        case "6657": //Enviar E-mail
                        case "6659": //Enviar Fax
                            csVariablesGlobales.FormularioEmail = csVariablesGlobales.SboApp.Forms.ActiveForm.UniqueID;
                            break;
                        case "1291":
                        case "1290":
                        case "1289":
                        case "1288":
                            if (csVariablesGlobales.SboApp.Forms.ActiveForm.TypeEx == "134")
                            {
                                csVariablesGlobales.PanelActualIC = csVariablesGlobales.SboApp.Forms.ActiveForm.PaneLevel;
                            }
                            break;
                        case "SubModelo347":
                            csFrmModelo347 FrmModelo347 = new csFrmModelo347();
                            //FrmModelo347.CargarFormulario();
                            Formulario = "FrmModelo347";
                            //Utilidades.LoadFromXML(csVariablesGlobales.StrPath + @"\XML\Formularios\FrmModelo347.xml", csVariablesGlobales.SboApp);
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\FrmModelo347.xml", csVariablesGlobales.SboApp);
                            FrmModelo347.CargarDatosInicialesDePantalla();
                            break;
                        case "SubMnuEntSalStock":
                            csFrmEntSalStock FrmEntSalStock = new csFrmEntSalStock();
                            //FrmEntSalStock.CargarFormulario();
                            Formulario = "FrmEntSalStock";
                            //Utilidades.LoadFromXML(csVariablesGlobales.StrPath + @"\XML\Formularios\FrmEntSalStock.xml", csVariablesGlobales.SboApp);
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\FrmEntSalStock.xml", csVariablesGlobales.SboApp);
                            FrmEntSalStock.CargarDatosInicialesDePantalla();
                            break;
                        case "SubMnuProcesos":
                            csFrmProcesos FrmProcesos = new csFrmProcesos();
                            //FrmProcesos.CargarFormulario();
                            Formulario = "FrmProcesos";
                            //Utilidades.LoadFromXML(csVariablesGlobales.StrPath + @"\XML\Formularios\FrmProcesos.xml", csVariablesGlobales.SboApp);
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\FrmProcesos.xml", csVariablesGlobales.SboApp);
                            FrmProcesos.CargarDatosInicialesDePantalla();
                            break;
                        case "SubModelo190":
                            csFrmModelo190 FrmModelo190 = new csFrmModelo190();
                            FrmModelo190.CargarFormulario();
                            Formulario = "FrmModelo190";
                            //LoadFromXML(csVariablesGlobales.StrPath + @"\XML\Formularios\FrmModelo190.xml", csVariablesGlobales.SboApp);
                            //LoadFromXML(@"\XML\Formularios\FrmModelo190.xml", csVariablesGlobales.SboApp);
                            break;
                        case "SubMnuGenMem":
                            csFrmGenerarMemoria FrmGenerarMemoria = new csFrmGenerarMemoria();
                            //FrmGenerarMemoria.CargarFormulario();
                            Formulario = "FrmGenerarMemoria";
                            //Utilidades.LoadFromXML(csVariablesGlobales.StrPath + @"\XML\Formularios\FrmGenerarMemoria.xml", csVariablesGlobales.SboApp);
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\FrmGenerarMemoria.xml", csVariablesGlobales.SboApp);
                            break;
                        case "SubModelo349":
                            csFrmModelo349 FrmModelo349 = new csFrmModelo349();
                            //FrmModelo349.CargarFormulario();
                            Formulario = "FrmModelo349";
                            //Utilidades.LoadFromXML(csVariablesGlobales.StrPath + @"\XML\Formularios\FrmModelo349.xml", csVariablesGlobales.SboApp);
                            csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\FrmModelo349.xml", csVariablesGlobales.SboApp);
                            FrmModelo349.CargarDatosInicialesDePantalla();
                            break;                        
                        case "1283":
                            oForm = csVariablesGlobales.SboApp.Forms.ActiveForm;
                            string Formu = csUtilidades.DameValor("[@SIA_FORMS]", "U_FormSIA", "U_FormSAP='" + oForm.Type.ToString() + "'");
                            switch (Formu)
                            {
                                case "@SIA_UBIC":
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                    if (pVal.BeforeAction == true)
                                    {
                                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_cod_ubic").Cells.Item(FilaSeleccionada).Specific;
                                        string Contador = csUtilidades.DameValor("OIBT", "Count(*)", "BatchNum = '" + oEditText.String + "'");
                                        if (Contador != "0")
                                        {
                                            csVariablesGlobales.SboApp.MessageBox("No se puede borrar porque hay o ha habido movimientos que hacen referencia a esta ubicación.", 1, "Ok", "", "");
                                            BubbleEvent = false;
                                        }
                                        FilaSeleccionada = 0;
                                    }
                                    break;
                            }
                            break;
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        case "1291":
                        case "1290":
                        case "1289":
                        case "1288":
                            if (csVariablesGlobales.SboApp.Forms.ActiveForm.TypeEx == "134" &&
                                csVariablesGlobales.PanelActualIC == 99)
                            {
                                oItem = oForm.Items.Item("fldCarProp");
                                oItem.Click(BoCellClickType.ct_Regular);
                            }
                            break;
                        case "1282": //Nuevo
                            if (csVariablesGlobales.SboApp.Forms.ActiveForm.TypeEx == "721" ||
                                csVariablesGlobales.SboApp.Forms.ActiveForm.TypeEx == "720")
                            {
                                oComboBox = (SAPbouiCOM.ComboBox)csVariablesGlobales.SboApp.Forms.ActiveForm.Items.Item("3").Specific;
                                oComboBox.Select("-2", BoSearchKey.psk_ByDescription);
                            }
                            break;
                        case "1287": //Duplicar
                            if (csVariablesGlobales.SboApp.Forms.ActiveForm.TypeEx == "721" ||
                                csVariablesGlobales.SboApp.Forms.ActiveForm.TypeEx == "720")
                            {
                                oComboBox = (SAPbouiCOM.ComboBox)csVariablesGlobales.SboApp.Forms.ActiveForm.Items.Item("3").Specific;
                                oComboBox.Select("-2", BoSearchKey.psk_ByDescription);
                            }
                            break;
                        case "6657": //Enviar E-mail
                        case "6659": //Enviar Fax
                            csVariablesGlobales.FormularioEnvioEmail = csVariablesGlobales.SboApp.Forms.ActiveForm.UniqueID;
                            Impresiones.Impresion(ref csVariablesGlobales.SboApp, ref pVal, out BubbleEvent);
                            if (BubbleEvent == false)
                            {
                                csVariablesGlobales.FormularioEmail = "";
                                csUtilidades.ActualizarDocumentoAEnviar();
                                return;
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Substring(0, 21) == "Form - already exists")
                {
                    oForm = csVariablesGlobales.SboApp.Forms.Item(Formulario);
                    oForm.Select();
                    return;
                }
                csVariablesGlobales.SboApp.MessageBox("ERROR - : " + ex.Message, 1, "Ok", "", "");
            }
        }

        private void SBOApp_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            //csUtilidades csUtilidades = new csUtilidades();
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    csUtilidades.ShutDown();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    csUtilidades.ShutDown();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    csUtilidades.ShutDown();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    csUtilidades.ShutDown();
                    break;
            }
        }

        private void SBOApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;           
            try
            {
                double Dto1, Dto2, Dto3, Dto4, Dto5, dblPrecio;
                string Formulario = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSIA", "U_FormSAP='" + pVal.FormTypeEx + "'");
                switch (Formulario)
                {
                    #region "133", "149", "139", "140", "180", "133", "179", "142", "143", "141", "181", "65308", "65300", "60090", "60091", "65309", "65301", "60092"
                    case "133":
                    case "149":
                    case "139":
                    case "140":
                    case "180":
                    case "179":
                    case "142":
                    case "143":
                    case "141":
                    case "181":
                    case "65308":
                    case "65300":
                    case "60090":
                    case "60091":
                    case "65309":
                    case "65301":
                    case "60092": //Documentos
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE &&
                            !pVal.BeforeAction)
                        {
                            csVariablesGlobales.InstanciaFormularioSAP = pVal.FormUID;
                        }
                        break;
                    #endregion
                    #region Seleccionar Report
                    case "SIASL00004":
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK & pVal.ItemUID == "matRep" && 
                            !pVal.BeforeAction)
                        {

                            csUtilidades.SBImprimirReportDoc(FormUID, ref pVal, "DocEntry");
                            csUtilidades.ActualizarDocumentoAEnviar();
                        }
                        break;
                    #endregion
                    #region SIA_REPORT
                    case "@SIA_REPORT":
                        if (pVal.EventType == BoEventTypes.et_CLICK)
                        {
                            if (pVal.ItemUID == "1" && (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.BeforeAction)
                            {
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                {

                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Report").Cells.Item(i).Specific;
                                    string Report = oEditText.String;
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_TipDoc").Cells.Item(i).Specific;
                                    string TipDoc = oEditText.String;
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(i).Specific;
                                    string Code = oEditText.String;
                                    if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_REPORT]", "Code", "Code='" + Code + "'") == "")
                                    {
                                        if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_REPORT]", "Code", "U_Report='" + Report + "' AND U_TipDoc='" + TipDoc + "'") == "" &&
                                            Report != "" && TipDoc != "")
                                        {
                                            string Valor = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(i).Specific;
                                            oEditText.Value = Valor;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Name").Cells.Item(i).Specific;
                                            oEditText.Value = Valor;
                                        }
                                        else
                                        {
                                            if (oMatrix.VisualRowCount <= oMatrix.RowCount)
                                            {
                                                BubbleEvent = true;
                                            }
                                            else
                                            {
                                                BubbleEvent = false;
                                                csVariablesGlobales.SboApp.MessageBox("El Report y el Tipo de Documento son obligatorios y no pueden estar repetidos", 1, "", "", "");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    #endregion
                    #region SIA_FORMS
                    case "@SIA_FORMS":
                        if (pVal.EventType == BoEventTypes.et_CLICK)
                        {
                            if (pVal.ItemUID == "1" && (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.BeforeAction)
                            {
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                {
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_formSAP").Cells.Item(i).Specific;
                                    string FormularioSAP = oEditText.String;
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_formSIA").Cells.Item(i).Specific;
                                    string FormularioSIA = oEditText.String;
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(i).Specific;
                                    string Code = oEditText.String;
                                    if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "Code", "Code='" + Code + "'") == "")
                                    {
                                        if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "Code", "U_formSAP='" + FormularioSAP + "' AND U_formSIA='" + FormularioSIA + "'") == "" &&
                                            FormularioSAP != "" && FormularioSIA != "")
                                        {
                                            string Valor = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(i).Specific;
                                            oEditText.Value = Valor;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Name").Cells.Item(i).Specific;
                                            oEditText.Value = Valor;
                                        }
                                        else
                                        {
                                            if (oMatrix.VisualRowCount <= oMatrix.RowCount)
                                            {
                                                BubbleEvent = true;
                                            }
                                            else
                                            {
                                                BubbleEvent = false;
                                                csVariablesGlobales.SboApp.MessageBox("El Formulario SAP y el Formulario SIA son obligatorios y no pueden estar repetidos", 1, "", "", "");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    #endregion
                    #region SIA_PARAM
                    case "@SIA_PARAM":
                        if (pVal.EventType == BoEventTypes.et_CLICK)
                        {
                            if (pVal.ItemUID == "1" && (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.BeforeAction)
                            {
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                {
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Concepto").Cells.Item(i).Specific;
                                    string Concepto = oEditText.String;
                                    oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(i).Specific;
                                    string Code = oEditText.String;
                                    if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "Code", "Code='" + Code + "'") == "")
                                    {
                                        if (csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "Code", "U_Concepto='" + Concepto + "'") == "" &&
                                            Concepto != "")
                                        {
                                            string Valor = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(i).Specific;
                                            oEditText.Value = Valor;
                                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Name").Cells.Item(i).Specific;
                                            oEditText.Value = Valor;
                                        }
                                        else
                                        {
                                            if (oMatrix.VisualRowCount <= oMatrix.RowCount)
                                            {
                                                BubbleEvent = true;
                                            }
                                            else
                                            {
                                                BubbleEvent = false;
                                                csVariablesGlobales.SboApp.MessageBox("El Concepto es obligatorio y no puede estar repetido", 1, "", "", "");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    #endregion
                    #region FrmImportacionesSIA
                    case "SIASL00001": //FrmImportacionesSIA
                        try
                        {
                            csFrmImportacionesSIA FrmImportacionesSIA = new csFrmImportacionesSIA();
                            FrmImportacionesSIA.FrmImportacionesSIA_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        }
                        catch { }
                        break;
                    #endregion
                    #region FrmImpresionListados
                    case "SIASL00002": //FrmImpresionListados
                        try
                        {
                            csFrmImpresionListados FrmImpresionListados = new csFrmImpresionListados();
                            FrmImpresionListados.FrmImpresionListados_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        }
                        catch { }
                        break;
                    #endregion
                    #region 134 - Interlocutores Comerciales
                    case "134":
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oForm.DataSources.UserDataSources.Add("dsFolder", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                            
                            oItem = oForm.Items.Add("fldCarProp", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                            oItem.Left = oForm.Items.Item("9").Left + oForm.Items.Item("9").Width;
                            oItem.Width = 100;
                            oItem.Top = oForm.Items.Item("9").Top;
                            oItem.Height = oForm.Items.Item("9").Height;
                            oItem.AffectsFormMode = false;
                            oFolder = ((SAPbouiCOM.Folder)(oItem.Specific));
                            oFolder.Caption = "Carac. Propias";
                            oFolder.DataBind.SetBound(true, "", "dsFolder");
                            oFolder.GroupWith("9");
                            #region Número Copias Facturas Venta
                            oItem = oForm.Items.Add("txtNCFV", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oItem.Left = oForm.Items.Item("43").Left;
                            oItem.Width = oForm.Items.Item("43").Width;
                            oItem.Top = oForm.Items.Item("43").Top;
                            oItem.Height = oForm.Items.Item("43").Height;
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNCFV").Specific;
                            oEditText.DataBind.SetBound(true, "OCRD", "U_NumCopFacVen");
                            oItem.FromPane = 99;
                            oItem.ToPane = 99;
                            oItem = oForm.Items.Add("lblNCFV", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            oItem.Left = oForm.Items.Item("44").Left;
                            oItem.Width = oForm.Items.Item("44").Width;
                            oItem.Top = oForm.Items.Item("44").Top;
                            oItem.Height = oForm.Items.Item("44").Height;
                            oItem.LinkTo = "txtNCFV";
                            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                            oStaticText.Caption = "Copias Factura Ventas";
                            oItem.FromPane = 99;
                            oItem.ToPane = 99;
                            #endregion
                            #region Número Copias Albarán Venta
                            oItem = oForm.Items.Add("txtNCAV", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oItem.Left = oForm.Items.Item("45").Left;
                            oItem.Width = oForm.Items.Item("45").Width;
                            oItem.Top = oForm.Items.Item("45").Top;
                            oItem.Height = oForm.Items.Item("45").Height;
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("txtNCAV").Specific;
                            oEditText.DataBind.SetBound(true, "OCRD", "U_NumCopAlbVen");
                            oItem.FromPane = 99;
                            oItem.ToPane = 99;
                            oItem = oForm.Items.Add("lblNCAV", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            oItem.Left = oForm.Items.Item("46").Left;
                            oItem.Width = oForm.Items.Item("46").Width;
                            oItem.Top = oForm.Items.Item("46").Top;
                            oItem.Height = oForm.Items.Item("46").Height;
                            oItem.LinkTo = "txtNCAV";
                            oStaticText = ((SAPbouiCOM.StaticText)(oItem.Specific));
                            oStaticText.Caption = "Copias Albarán Ventas";
                            oItem.FromPane = 99;
                            oItem.ToPane = 99;
                            #endregion
                        }
                        if (pVal.EventType == BoEventTypes.et_CLICK && !pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "fldCarProp")
                            {
                                oForm.PaneLevel = 99;
                            }
                        }
                        break;
                    #endregion
                    #region 392
                    case "392":
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oItem = oForm.Items.Add("btnAsiGas", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 20;
                            oItem.Width = oForm.Items.Item("2").Width + 10;
                            oItem.Top = oForm.Items.Item("2").Top;
                            oItem.FontSize = oForm.Items.Item("2").FontSize;
                            oItem.Height = oForm.Items.Item("2").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnAsiGas").Specific;
                            oButton.Caption = "Asiento Gastos";
                            oItem.Visible = true;
                            oItem = oForm.Items.Add("btnCarRef", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 100;
                            oItem.Width = oForm.Items.Item("2").Width + 10;
                            oItem.Top = oForm.Items.Item("2").Top;
                            oItem.FontSize = oForm.Items.Item("2").FontSize;
                            oItem.Height = oForm.Items.Item("2").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnCarRef").Specific;
                            oButton.Caption = "Cargar Referencia";
                            oItem.Visible = true;
                        }
                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "btnAsiGas")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    csFrmAsientoGastos FrmAsientoGastos = new csFrmAsientoGastos();
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("5").Specific;
                                    csVariablesGlobales.NumeroAsiento = Convert.ToInt32(oEditText.String);
                                    csVariablesGlobales.FormularioAsientoGastosAbierto = true;
                                    Formulario = "FrmAsientoGastos";
                                    csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath + @"\Formularios\FrmAsientoGastos.xml", csVariablesGlobales.SboApp);
                                    FrmAsientoGastos.CargarDatosInicialesDePantalla();                                    
                                    break;
                                }
                                else
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Actualiace primero antes de continuar con este proceso.", 1, "Ok", "", "");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "btnCarRef")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    SAPbobsCOM.JournalEntries oJournalEntries;
                                    oJournalEntries = (SAPbobsCOM.JournalEntries)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries));
                                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("5").Specific;
                                    oJournalEntries.GetByKey(Convert.ToInt32(oEditText.String));
                                    string Valor = "";
                                    for (int i = 0; i <= oJournalEntries.Lines.Count - 1; i++)
                                    {
                                        oJournalEntries.Lines.SetCurrentLine(i);                                        
                                        Valor = csUtilidades.DameValor("OBOE", "RefNum", "BoeKey =" + 
                                        oJournalEntries.Lines.Reference1);
                                        if (Valor != "")
                                            {
                                                Valor = oJournalEntries.Lines.LineMemo + "  -  " + Valor;
                                                Valor = Valor.PadRight(50, ' ').Substring(0, 49).Trim();
                                                oJournalEntries.Lines.LineMemo = Valor;
                                                //csVariablesGlobales.SboApp.MessageBox(oJournalEntries.Lines.LineMemo, 1, "Ok", "", "");
                                            }                                        
                                    }
                                    int nRet;
                                    nRet = oJournalEntries.Update();
                                    if (nRet != 0)
                                    {
                                        csVariablesGlobales.SboApp.SetStatusBarMessage(csVariablesGlobales.oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short, true);
                                    }
                                    else
                                    {                                    
                                      csVariablesGlobales.SboApp.MessageBox("La carga se ha realizado correctamente.", 1, "Ok", "", "");
                                    }
                                    //oForm.Close();
                                    break;
                                }
                                else
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Actualiace primero antes de continuar con este proceso.", 1, "Ok", "", "");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }

                        if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction &&
                            csVariablesGlobales.NumeroAsiento != 0 &&
                            !csVariablesGlobales.FormularioAsientoGastosAbierto)
                        {
                            csVariablesGlobales.NumeroAsiento = 0;
                            csVariablesGlobales.SboApp.ActivateMenuItem("1291");
                        }
                        break;
                    #endregion
                    #region FrmAsientoGastos
                    case "SIASL00003": //FrmAsientoGastos
                        csFrmAsientoGastos FrmAsientoGastos2 = new csFrmAsientoGastos();
                        FrmAsientoGastos2.FrmAsientoGastos_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion
                    #region 606
                    case "606": //Depósitos
                        #region Crear Botones
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oItem = oForm.Items.Add("btnNorma58", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 20;
                            oItem.Top = oForm.Items.Item("2").Top;
                            oItem.FontSize = oForm.Items.Item("2").FontSize;
                            oItem.Height = oForm.Items.Item("2").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnNorma58").Specific;
                            oButton.Caption = "Norma 58";
                            oItem.Visible = true;
                            oItem = oForm.Items.Add("btnNorma19", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("btnNorma58").Left + oForm.Items.Item("btnNorma58").Width + 20;
                            oItem.Top = oForm.Items.Item("btnNorma58").Top;
                            oItem.FontSize = oForm.Items.Item("btnNorma58").FontSize;
                            oItem.Height = oForm.Items.Item("btnNorma58").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnNorma19").Specific;
                            oButton.Caption = "Norma 19";
                            oItem.Visible = true;
                        }
                        #endregion
                        #region Norma 58
                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnNorma58" && pVal.BeforeAction)
                        {
                            string Ruta = "";
                            csSaveFileDialog SaveFileDialog = new csSaveFileDialog();
                            SaveFileDialog.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                            //string DirectorioActual = Environment.CurrentDirectory;
                            SaveFileDialog.InitialDirectory = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros Modelos'");
                            Thread threadGetFile = new Thread(new ThreadStart(SaveFileDialog.GetFileName));
                            threadGetFile.TrySetApartmentState(ApartmentState.STA);
                            try
                            {
                                threadGetFile.Start();
                                while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                                Thread.Sleep(1);  // Wait a sec more
                                threadGetFile.Join();    // Wait for thread to end

                                // Use file name as you will here
                                Ruta = SaveFileDialog.FileName;
                            }
                            catch (Exception ex)
                            {
                                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                            }
                            threadGetFile = null;
                            SaveFileDialog = null;
                            if (Ruta != "")
                            {
                                csNormas Normas = new csNormas();
                                Normas.GenerarFicheroNorma58(Ruta, oForm);
                            }
                            else
                            {
                                csVariablesGlobales.SboApp.MessageBox("Se debe indicar un fichero obligatoriamente", 1, "OK", "", "");
                            }
                            //Environment.CurrentDirectory = DirectorioActual;
                        }
                        #endregion
                        #region Norma 19
                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "btnNorma19" && pVal.BeforeAction)
                        {
                            string Ruta = "";
                            csSaveFileDialog SaveFileDialog = new csSaveFileDialog();
                            //string DirectorioActual = Environment.CurrentDirectory;
                            SaveFileDialog.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                            SaveFileDialog.InitialDirectory = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros Modelos'");
                            Thread threadGetFile = new Thread(new ThreadStart(SaveFileDialog.GetFileName));
                            threadGetFile.TrySetApartmentState(ApartmentState.STA);
                            try
                            {
                                threadGetFile.Start();
                                while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                                Thread.Sleep(1);  // Wait a sec more
                                threadGetFile.Join();    // Wait for thread to end

                                // Use file name as you will here
                                Ruta = SaveFileDialog.FileName;
                            }
                            catch (Exception ex)
                            {
                                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                            }
                            threadGetFile = null;
                            SaveFileDialog = null;
                            if (Ruta != "")
                            {
                                csNormas Normas2 = new csNormas();
                                Normas2.GeneroFicheroNorma19(Ruta, oForm);
                            }
                            else
                            {
                                csVariablesGlobales.SboApp.MessageBox("Se debe indicar un fichero obligatoriamente", 1, "OK", "", "");
                            }
                            //Environment.CurrentDirectory = DirectorioActual;
                        }
                        #endregion
                        break;
                    #endregion
                    #region - FrmMensaje
                    case "SIASL00005": //FrmMensaje
                        csFrmMensaje FrmMensaje = new csFrmMensaje();
                        FrmMensaje.FrmMensaje_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion
                    #region 3002 - Documentos Preliminares
                    case "3002": //Documentos Preliminares
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oItem = oForm.Items.Add("btnBorrar", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 20;
                            oItem.Width = oForm.Items.Item("2").Width + 30;
                            oItem.Top = oForm.Items.Item("2").Top;
                            oItem.FontSize = oForm.Items.Item("2").FontSize;
                            oItem.Height = oForm.Items.Item("2").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnBorrar").Specific;
                            oButton.Caption = "Borrar Documentos";
                            oItem.Visible = true;
                        }
                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "btnBorrar")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    //csFrmMensaje FrmMensaje = new csFrmMensaje();
                                    //FrmMensaje.CargarFormulario();
                                    Formulario = "FrmMensaje";
                                    csUtilidades.LoadFromXML(System.Windows.Forms.Application.StartupPath +
                                                           @"\Formularios\FrmMensaje.xml",
                                                           csVariablesGlobales.SboApp);
                                    csVariablesGlobales.FormularioDesde = "3002";
                                    csVariablesGlobales.FormularioDesdeUID = FormUID;
                                    break;
                                }
                                else
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Actualiace primero antes de continuar con este proceso.", 1, "Ok", "", "");
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        if (pVal.EventType == BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction &&
                            csVariablesGlobales.NumeroAsiento != 0 &&
                            !csVariablesGlobales.FormularioAsientoGastosAbierto)
                        {
                            csVariablesGlobales.NumeroAsiento = 0;
                            csVariablesGlobales.SboApp.ActivateMenuItem("1291");
                        }
                        break;
                    #endregion
                    #region @SIA_PLANO
                    case "@SIA_PLANO": //Planos
                        string Articulo;
                        if (pVal.ColUID == "U_cod_artic" && pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            if ((SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_ADD_MODE || (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                SAPbouiCOM.EditText oTxt = null;
                                SAPbouiCOM.Matrix oMatrix = null;
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                //oForm.Close()
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_cod_artic").Cells.Item(pVal.Row).Specific;
                                Articulo = oTxt.Value;
                                if (Articulo == "")
                                {
                                    return;
                                }
                                if (csUtilidades.DameValor("[OITM]", "ItemCode", "ItemCode='" + Articulo + "'") == "")
                                {
                                    csVariablesGlobales.SboApp.MessageBox("No existe ningún artículo con ese código", 1, "Ok", "", "");
                                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_cod_artic").Cells.Item(pVal.Row).Specific;
                                    oTxt.Value = "";
                                    oTxt.Active = true;
                                    return;
                                }
                                if (csUtilidades.DameValor("[@SIA_PLANO]", "U_cod_artic", "U_cod_artic='" + Articulo + "'") != "")
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Ya existe una línea con ese código de artículo ", 1, "Ok", "", "");
                                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_cod_artic").Cells.Item(pVal.Row).Specific;
                                    oTxt.Value = "";
                                    oTxt.Active = true;
                                    return;
                                }
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific;
                                string UltimoValor = csUtilidades.UltimoCode("[@SIA_PLANO]");
                                oTxt.Value = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Name").Cells.Item(pVal.Row).Specific;
                                oTxt.Value = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                            }
                        }

                        break;
                    #endregion
                    #region @SIA_UBIC
                    case "@SIA_UBIC": //Ubicaciones
                        string Ubicacion;
                        if (pVal.ColUID == "U_cod_ubic" && pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            if ((SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_ADD_MODE || (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                SAPbouiCOM.EditText oTxt = null;
                                SAPbouiCOM.Matrix oMatrix = null;
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                //oForm.Close()
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_cod_ubic").Cells.Item(pVal.Row).Specific;
                                if (!csUtilidades.TodasMayusculas(oTxt.String))
                                {
                                    oTxt.String = oTxt.String.ToUpper();
                                }
                                Ubicacion = oTxt.Value;
                                if (Ubicacion == "")
                                {
                                    return;
                                }
                                if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "'") != "")
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Ya existe una ubicación con ese código", 1, "Ok", "", "");
                                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_cod_ubic").Cells.Item(pVal.Row).Specific;
                                    oTxt.Value = "";
                                    oTxt.Active = true;
                                    BubbleEvent = false;
                                    return;
                                }
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific;
                                string UltimoValor = csUtilidades.UltimoCode("[@SIA_UBIC]");
                                oTxt.Value = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Name").Cells.Item(pVal.Row).Specific;
                                oTxt.Value = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                            }
                        }

                        break;
                    #endregion
                    #region 41
                    case "41":
                        if (pVal.ColUID == "2" & pVal.EventType == BoEventTypes.et_LOST_FOCUS & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                            if ((SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_ADD_MODE ||
                                (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                int i = 0;
                                SAPbouiCOM.Matrix oMatrix;
                                SAPbouiCOM.EditText oTxt;
                                string Almacen;
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("35").Specific;
                                for (i = 1; i <= oMatrix.RowCount; i++)
                                {
                                    if (oMatrix.IsRowSelected(i))
                                    {
                                        break;
                                    }
                                }
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("40").Cells.Item(i).Specific;
                                Almacen = oTxt.Value;
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                                if (!csUtilidades.TodasMayusculas(oTxt.String))
                                {
                                    oTxt.String = oTxt.String.ToUpper();
                                }
                                Ubicacion = oTxt.Value;
                                if (Ubicacion == "")
                                {
                                    return;
                                }

                                if (csUtilidades.DameValor("[@SIA_UBIC]", "U_cod_ubic", "U_cod_ubic='" + Ubicacion + "' AND U_almacen='" + Almacen + "'") == "")
                                {
                                    csVariablesGlobales.SboApp.MessageBox("No existe ninguna ubicación con ese código para ese almacén", 1, "Ok", "", "");
                                    oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                                    oTxt.Value = "";
                                    return;
                                }
                            }
                        }

                        if (pVal.ColUID != "2" & pVal.EventType == BoEventTypes.et_GOT_FOCUS & pVal.BeforeAction == false & pVal.ActionSuccess == true)
                        {
                            SAPbouiCOM.Matrix oMatrix;
                            SAPbouiCOM.EditText oTxt;
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                            oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific;
                            Ubicacion = oTxt.Value;
                            if (Ubicacion == "")
                            {
                                oTxt.Active = true;
                                return;
                            }
                        }
                        break;
                    #endregion
                    #region FrmEntSalStock
                    case "1060001": //FrmEntSalStock
                        csFrmEntSalStock FrmEntSalStock = new csFrmEntSalStock();
                        FrmEntSalStock.FrmEntSalStock_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion
                    #region FrmProcesos
                    case "1060003": //FrmProcesos
                        csFrmProcesos FrmProcesos = new csFrmProcesos();
                        FrmProcesos.FrmProcesos_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion
                    #region FrmModelo347
                    case "1060009": //FrmModelo347
                        csFrmModelo347 FrmModelo347 = new csFrmModelo347();
                        FrmModelo347.FrmModelo347_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion
                    #region FrmModelo349
                    case "1060013": //FrmModelo349
                        csFrmModelo349 FrmModelo349 = new csFrmModelo349();
                        FrmModelo349.FrmModelo349_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion
                    #region FrmGenerarMemoria
                    case "1060010": //FrmGenerarMemoria
                        csFrmGenerarMemoria FrmGenerarMemoria = new csFrmGenerarMemoria();
                        FrmGenerarMemoria.FrmGenerarMemoria_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                        break;
                    #endregion                    
                    #region @SIA_PEDC
                    case "@SIA_PEDC": //Reports
                        if (pVal.ColUID == "U_SIA_NumPed" && pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.BeforeAction == false && pVal.ActionSuccess == true)
                        {
                            if ((SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_ADD_MODE || (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                SAPbouiCOM.EditText oTxt = null;
                                SAPbouiCOM.Matrix oMatrix = null;
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Code").Cells.Item(pVal.Row).Specific;
                                string UltimoValor = csUtilidades.UltimoCode("[@SIA_PEDC]");
                                oTxt.Value = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                                oTxt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("Name").Cells.Item(pVal.Row).Specific;
                                oTxt.Value = csUtilidades.CompletaConCeros(8, csUtilidades.UltimoCode("[" + Formulario + "]"), 1);
                            }
                        }

                        break;
                    #endregion
                    #region 65010 - 347
                    case "65010":
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oItem = oForm.Items.Add("btnFich", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("1").Left;
                            oItem.Width = oForm.Items.Item("1").Width;
                            oItem.Top = oForm.Items.Item("1").Top - 22;
                            oItem.FontSize = oForm.Items.Item("1").FontSize;
                            oItem.Height = oForm.Items.Item("1").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnFich").Specific;
                            oButton.Caption = "Fichero SIA";
                            oItem.Visible = true;                            
                                                    
                            oItem = oForm.Items.Add("btnCarta", SAPbouiCOM.BoFormItemTypes.it_BUTTON);                            
                            oItem.Left = oForm.Items.Item("4").Left;
                            oItem.Width = oForm.Items.Item("4").Width;
                            oItem.Top = oForm.Items.Item("4").Top - 22;
                            oItem.Height = oForm.Items.Item("4").Height;

                            oItem.FontSize = oForm.Items.Item("1").FontSize;                            
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnCarta").Specific;
                            oButton.Caption = "Carta 347";
                            oItem.Visible = true;
                        }
                        if (pVal.EventType == BoEventTypes.et_CLICK && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            if (pVal.ItemUID == "btnFich")
                            {
                                string Ruta347 = "";                                
                                csSaveFileDialog SaveFileDialog = new csSaveFileDialog();
                                SaveFileDialog.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                                SaveFileDialog.InitialDirectory = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Ruta", "U_TipoRuta = 'Destino Ficheros Modelos En Local'");
                                Thread threadGetFile = new Thread(new ThreadStart(SaveFileDialog.GetFileName));
                                threadGetFile.TrySetApartmentState(ApartmentState.STA);
                                try
                                {
                                    threadGetFile.Start();
                                    while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                                    Thread.Sleep(1);  // Wait a sec more
                                    threadGetFile.Join();    // Wait for thread to end

                                    // Use file name as you will here
                                    Ruta347 = SaveFileDialog.FileName;
                                }
                                catch (Exception ex)
                                {
                                    csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "OK", "", "");
                                }
                                threadGetFile = null;
                                SaveFileDialog = null;
                                if (Ruta347 != "")
                                {
                                    csModelos.GenerarFichero347(Ruta347, oForm);
                                }
                                else
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Se debe indicar un fichero obligatoriamente", 1, "OK", "", "");
                                }                             
                            }

                          if (pVal.ItemUID == "btnCarta")
                            {
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);                         
                                string Archivo = "Informe347";
                                string StrNifCli = "", StrNomCli = "", cNumFac = "", cNIFOK = "", cTabla = "", cAuxiliar = "";
                                string StrCodCli = "", StrTipoIC = "", cFecha = "", cSiguiente = "";
                                string Suma = "";
                                int nMes = 0;
                                double PrimeTri = 0, SegunTri = 0, TercerTri = 0, CuartoTri = 0;
                                SAPbobsCOM.Recordset oRecordsetLis;
                                SAPbobsCOM.Recordset oRecordsetFech;                                
                                oRecordsetLis = ((SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
                                oRecordsetFech = ((SAPbobsCOM.Recordset)csVariablesGlobales.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset));
                              

                                csImpresiones Impresiones = new csImpresiones();
                                #region Generar Fichero XML
                                XmlDocument oXML = new XmlDocument();
                                XmlNode oBase, oTabla;
                                XmlAttribute[] oCampo = new XmlAttribute[15];
                                //Creao la base de datos            
                                oBase = oXML.CreateElement("Informe347");
                                oXML.AppendChild(oBase);

                                //Para la tabla
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                                {                    
                                       
                                  StrNifCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(i).Specific).String;
                                  cNumFac = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("350000012").Cells.Item(i).Specific).String;

                                        //Si esta en blanco o no hay nada, continue                      
                                  if ( (StrNifCli == "" && cNumFac == "") || (cNumFac != "" && cNIFOK == "") ) continue;


                                  if (StrNifCli != "")
                                  {
                                    cNIFOK = StrNifCli;
                                    StrCodCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("0").Cells.Item(i).Specific).String;
                                    StrNomCli = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).String;                                      
                                    oRecordsetLis.DoQuery("SELECT CardType FROM OCRD WHERE OCRD.CardCode = '" + StrCodCli + "'");
                                    oRecordsetLis.MoveFirst();
                                    StrTipoIC = oRecordsetLis.Fields.Item("CardType").Value.ToString();
                                    Suma = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                                   }


                                    if (cNumFac != "")
                                    {
                                      cTabla = "";
                                      if (cNumFac.Replace("FA", "") != cNumFac)
                                        {
                                          cAuxiliar = cNumFac.Replace("FA", "");
                                          cTabla = "OINV";
                                        }
                                       else if (cNumFac.Replace("RC", "") != cNumFac)
                                        {
                                          cAuxiliar = cNumFac.Replace("RC", "");
                                          cTabla = "ORIN";
                                        }
                                        else if (cNumFac.Replace("AN", "") != cNumFac)
                                        {                                               
                                          cAuxiliar = cNumFac.Replace("AN", "");
                                          if (StrTipoIC == "C")
                                            {
                                              cTabla = "ODPI";
                                            }
                                          else
                                            {
                                              cTabla = "ODPO";
                                            }
                                         }
                                         else if (cNumFac.Replace("TT", "") != cNumFac)
                                            {
                                             cAuxiliar = cNumFac.Replace("TT", "");
                                              cTabla = "OPCH";
                                         }
                                         else if (cNumFac.Replace("AC", "") != cNumFac)
                                         {
                                           cAuxiliar = cNumFac.Replace("AC", "");
                                           cTabla = "ORPC";
                                         }
                                         else
                                         {
                                           csVariablesGlobales.SboApp.MessageBox("Error en 347 para documento" + cNumFac, 0, "", "", "");
                                           cTabla = "";
                                           return;
                                         }


                                         oRecordsetFech.DoQuery("SELECT TaxDate FROM " + cTabla + " WHERE DocNum = " + cAuxiliar);
                                         oRecordsetFech.MoveFirst();
                                         cFecha = oRecordsetFech.Fields.Item("TaxDate").Value.ToString();


                                          nMes = Convert.ToDateTime(cFecha).Month;
                                          if (nMes == 1 || nMes == 2 || nMes == 3)
                                          {
                                                cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                                                cAuxiliar = cAuxiliar.Replace(",", "#");
                                                cAuxiliar = cAuxiliar.Replace(".", ",");
                                                cAuxiliar = cAuxiliar.Replace("#", ".");
                                                PrimeTri = PrimeTri + Convert.ToDouble(cAuxiliar);
                                          }
                                          else if (nMes == 4 || nMes == 5 || nMes == 6)
                                            {
                                                cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                                                cAuxiliar = cAuxiliar.Replace(",", "#");
                                                cAuxiliar = cAuxiliar.Replace(".", ",");
                                                cAuxiliar = cAuxiliar.Replace("#", ".");
                                                SegunTri = SegunTri + Convert.ToDouble(cAuxiliar);
                                            }
                                          else if (nMes == 7 || nMes == 8 || nMes == 9)
                                            {
                                                cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                                                cAuxiliar = cAuxiliar.Replace(",", "#");
                                                cAuxiliar = cAuxiliar.Replace(".", ",");
                                                cAuxiliar = cAuxiliar.Replace("#", ".");
                                                TercerTri = TercerTri + Convert.ToDouble(cAuxiliar);
                                            }
                                          else if (nMes == 10 || nMes == 11 || nMes == 12)
                                            {
                                                cAuxiliar = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(i).Specific).String;
                                                cAuxiliar = cAuxiliar.Replace(",", "#");
                                                cAuxiliar = cAuxiliar.Replace(".", ",");
                                                cAuxiliar = cAuxiliar.Replace("#", ".");
                                                CuartoTri = CuartoTri + Convert.ToDouble(cAuxiliar);
                                            }
                                        }

                                        //Si estoy en el ultimo, pinto
                                        if (i < oMatrix.VisualRowCount)
                                        {
                                            cSiguiente = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(i + 1).Specific).String;
                                        }

                                        if (i == oMatrix.VisualRowCount || (cSiguiente != "" && cSiguiente != cNIFOK))
                                        {
                                            oTabla = oXML.CreateElement("Informe347");
                                            oCampo[1] = oXML.CreateAttribute("NomIC");
                                            oCampo[2] = oXML.CreateAttribute("Street");
                                            oCampo[3] = oXML.CreateAttribute("City");
                                            oCampo[4] = oXML.CreateAttribute("ZipCode");
                                            oCampo[5] = oXML.CreateAttribute("County");
                                            oCampo[6] = oXML.CreateAttribute("Importe");
                                            oCampo[7] = oXML.CreateAttribute("Ejercicio");
                                            oCampo[8] = oXML.CreateAttribute("Total");
                                            oCampo[9] = oXML.CreateAttribute("Trimestre1");
                                            oCampo[10] = oXML.CreateAttribute("Trimestre2");
                                            oCampo[11] = oXML.CreateAttribute("Trimestre3");
                                            oCampo[12] = oXML.CreateAttribute("Trimestre4");
                                            oCampo[13] = oXML.CreateAttribute("Empresa");


                                            oCampo[1].InnerText = StrNomCli;
                                            oCampo[2].InnerText = csUtilidades.DameValor("OCRD", "Address", "CardCode ='" + StrCodCli + "'");
                                            oCampo[3].InnerText = csUtilidades.DameValor("OCRD", "City", "CardCode ='" + StrCodCli + "'");
                                            oCampo[4].InnerText = csUtilidades.DameValor("OCRD", "ZipCode", "CardCode ='" + StrCodCli + "'");
                                            oCampo[5].InnerText = csUtilidades.DameValor("OCRD", "County", "CardCode ='" + StrCodCli + "'");
                                            oCampo[6].InnerText = csUtilidades.DameValor("OADM", "MinAmnt347", "");
                                            oCampo[7].InnerText = csVariablesGlobales.AñoModelo;
                                            oCampo[8].InnerText = Suma;
                                            oCampo[9].InnerText = Convert.ToString(PrimeTri);
                                            oCampo[10].InnerText = Convert.ToString(SegunTri);
                                            oCampo[11].InnerText = Convert.ToString(TercerTri);
                                            oCampo[12].InnerText = Convert.ToString(CuartoTri);
                                            oCampo[13].InnerText = csUtilidades.DameValor("OADM", "CompnyName", "");                                                                                      

                                            oTabla.Attributes.Append(oCampo[1]);
                                            oTabla.Attributes.Append(oCampo[2]);
                                            oTabla.Attributes.Append(oCampo[3]);
                                            oTabla.Attributes.Append(oCampo[4]);
                                            oTabla.Attributes.Append(oCampo[5]);
                                            oTabla.Attributes.Append(oCampo[6]);
                                            oTabla.Attributes.Append(oCampo[7]);
                                            oTabla.Attributes.Append(oCampo[8]);
                                            oTabla.Attributes.Append(oCampo[9]);
                                            oTabla.Attributes.Append(oCampo[10]);
                                            oTabla.Attributes.Append(oCampo[11]);
                                            oTabla.Attributes.Append(oCampo[12]);
                                            oTabla.Attributes.Append(oCampo[13]);
                                            oXML.DocumentElement.AppendChild(oTabla);

                                            csVariablesGlobales.SboApp.SetStatusBarMessage("Generada Carta para el cliente " + StrCodCli + " NIF " + cNIFOK, BoMessageTime.bmt_Short, false);
                                            PrimeTri = 0; SegunTri = 0; TercerTri = 0; CuartoTri = 0;
                                            Suma = "";                                            
                                        }
                                    }                                
                                oXML.Save(csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]",
                                              "U_Ruta", "U_TipoRuta = 'Destino Ficheros Temporales En Local'") + Archivo + ".xml");
                                oXML = null;
                                #endregion
                                Impresiones.Informe(csVariablesGlobales.StrRutRep + Archivo + ".rpt", "");
                            }
                        }
                        break;
                    #endregion
                    #region 65011
                    case "65011":
                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction && !pVal.ActionSuccess)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "9")
                            {
                                string Año = ((SAPbouiCOM.EditText)oForm.Items.Item("5").Specific).String;
                                csVariablesGlobales.AñoModelo = Año.Substring(Año.Length - 4, 4);
                            }
                        }
                        break;
                    #endregion
                    #region 60051 Remesas
                    case "60051":
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oItem = oForm.Items.Add("btnImpSel", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("100").Left - oForm.Items.Item("1").Width - 20;
                            oItem.Width = oForm.Items.Item("100").Width + 10;
                            oItem.Top = oForm.Items.Item("100").Top;
                            oItem.FontSize = oForm.Items.Item("100").FontSize;
                            oItem.Height = oForm.Items.Item("100").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnImpSel").Specific;
                            oButton.Caption = "Imp. Selección";
                            oItem.Visible = true;
                        }

                        if (pVal.EventType == BoEventTypes.et_CLICK && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            bool ExistenLineas = false;
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "btnImpSel")
                            {
                                string Archivo = "CartaRemesa";
                                csImpresiones Impresiones = new csImpresiones();
                                #region Generar Fichero XML 
                                XmlDocument oXML = new XmlDocument();
                                XmlNode oBase, oTabla;
                                XmlAttribute[] oCampo = new XmlAttribute[15];
                                //Creao la base de datos            
                                oBase = oXML.CreateElement("BaseXML");
                                oXML.AppendChild(oBase);

                                //Tabla con el banco
                                string BankCode = ((SAPbouiCOM.ComboBox)oForm.Items.Item("46").Specific).Selected.Value;
                                string Country = ((SAPbouiCOM.ComboBox)oForm.Items.Item("45").Specific).Selected.Value;
                                string Account = ((SAPbouiCOM.ComboBox)oForm.Items.Item("49").Specific).Selected.Value;
                                oTabla = oXML.CreateElement("Banco");
                                oCampo[1] = oXML.CreateAttribute("NombreBanco");
                                oCampo[2] = oXML.CreateAttribute("CalleBanco");
                                oCampo[3] = oXML.CreateAttribute("CPBanco");
                                oCampo[4] = oXML.CreateAttribute("CiudadBanco");
                                oCampo[5] = oXML.CreateAttribute("ProvinBanco");
                                oCampo[6] = oXML.CreateAttribute("AttBanco");
                                oCampo[7] = oXML.CreateAttribute("Cuenta");
                                oCampo[1].InnerText = csUtilidades.DameValor("ODSC", "BankName", "BankCode ='" + BankCode + "' AND CountryCod = '" + Country + "'");
                                oCampo[2].InnerText = csUtilidades.DameValor("DSC1", "Street", "BankCode ='" + BankCode + "' AND Country = '" + Country + "' AND Account ='" + Account + "'");
                                oCampo[3].InnerText = csUtilidades.DameValor("DSC1", "ZipCode", "BankCode ='" + BankCode + "' AND Country = '" + Country + "' AND Account ='" + Account + "'");
                                oCampo[4].InnerText = csUtilidades.DameValor("DSC1", "City", "BankCode ='" + BankCode + "' AND Country = '" + Country + "' AND Account ='" + Account + "'");
                                oCampo[5].InnerText = csUtilidades.DameValor("DSC1", "County", "BankCode ='" + BankCode + "' AND Country = '" + Country + "' AND Account ='" + Account + "'");
                                oCampo[6].InnerText = "";
                                oCampo[7].InnerText = Account;
                                oTabla.Attributes.Append(oCampo[1]);
                                oTabla.Attributes.Append(oCampo[2]);
                                oTabla.Attributes.Append(oCampo[3]);
                                oTabla.Attributes.Append(oCampo[4]);
                                oTabla.Attributes.Append(oCampo[5]);
                                oTabla.Attributes.Append(oCampo[6]);
                                oTabla.Attributes.Append(oCampo[7]);
                                oXML.DocumentElement.AppendChild(oTabla);

                                //Para la tabla
                                oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("5").Specific;
                                for (int nCont = 1; nCont <= oMatrix.VisualRowCount; nCont++)
                                {
                                    oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("9").Cells.Item(nCont).Specific;
                                    if (oCheckBox.Checked)
                                    {
                                        oTabla = oXML.CreateElement("Efectos");
                                        oCampo[1] = oXML.CreateAttribute("CardName");
                                        oCampo[2] = oXML.CreateAttribute("Importe");
                                        oCampo[3] = oXML.CreateAttribute("ViaPago");
                                        oCampo[4] = oXML.CreateAttribute("FechaVto");
                                        oCampo[5] = oXML.CreateAttribute("FechaRemesa");
                                        oCampo[6] = oXML.CreateAttribute("DescripViaPago");
                                        oCampo[1].InnerText = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("3").Cells.Item(nCont).Specific).String;
                                        oCampo[2].InnerText = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("2").Cells.Item(nCont).Specific).String;
                                        oCampo[3].InnerText = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("13").Cells.Item(nCont).Specific).String;
                                        oCampo[4].InnerText = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("17").Cells.Item(nCont).Specific).String;
                                        oCampo[5].InnerText = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("4").Cells.Item(nCont).Specific).String;
                                        oCampo[6].InnerText = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("12").Cells.Item(nCont).Specific).String;
                                        oTabla.Attributes.Append(oCampo[1]);
                                        oTabla.Attributes.Append(oCampo[2]);
                                        oTabla.Attributes.Append(oCampo[3]);
                                        oTabla.Attributes.Append(oCampo[4]);
                                        oTabla.Attributes.Append(oCampo[5]);
                                        oTabla.Attributes.Append(oCampo[6]);
                                        oXML.DocumentElement.AppendChild(oTabla);
                                        ExistenLineas = true;
                                    }
                                }
                                oXML.Save(csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]",
                                              "U_Ruta", "U_TipoRuta = 'Destino Ficheros Temporales En Local'") + Archivo + ".xml");
                                #endregion
                                if (ExistenLineas)
                                {
                                    Impresiones.Informe(csVariablesGlobales.StrRutRep + Archivo + ".rpt", "");
                                }
                                else
                                {
                                    csVariablesGlobales.SboApp.MessageBox("No hay lineas seleccionadas", 1, "Ok", "", "");
                                }
                            }
                            //if (pVal.ItemUID == "37" || pVal.ItemUID == "38")
                            //{
                            //    csVariablesGlobales.ParaCambiodeReferencia = true;
                            //}
                            //if (pVal.ItemUID == "15" || pVal.ItemUID == "39" || 
                            //    pVal.ItemUID == "34" || pVal.ItemUID == "36")
                            //{
                            //    csVariablesGlobales.ParaCambiodeReferencia = false;
                            //}
                        }
                        //if (pVal.ItemUID == "1" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        //{
                        //    string Valor = ((SAPbouiCOM.ComboBox)oForm.Items.Item("4").Specific).Selected.Value;
                        //    if (Valor == "D" || Valor == "P")
                        //    {
                                
                        //    }
                        //}
                        //if (pVal.EventType == BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                        //{
                        //    csVariablesGlobales.ParaCambiodeReferencia = false;
                        //}
                        break;
                    #endregion
                    #region 720 - Salida Mercancías
                    case "720":
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("3").Specific;
                            oComboBox.Select("-2", BoSearchKey.psk_ByDescription);
                        }
                        break;
                    #endregion
                    #region 721 - Entrada Mercancías
                    case "721":
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("3").Specific;
                            oComboBox.Select("-2", BoSearchKey.psk_ByDescription);
                        }
                        break;
                    #endregion
                }
                switch (Formulario)
                {
                    #region 133
                    case "133":
                        if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction == true &&
                            pVal.ItemUID == "47")
                        {
                            csVariablesGlobales.SboApp.MessageBox("La condición de pago ha cambiado. Revise el descuento comercial.", 1, "Ok", "", "");
                        }
                        break;
                    #endregion
                    #region 139
                    case "139":

                        if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction == true &&
                            pVal.ItemUID == "47")
                        {
                            csVariablesGlobales.SboApp.MessageBox("La condición de pago ha cambiado. Revise la lista de precios.", 1, "Ok", "", "");
                        }

                        break;
                    #endregion
                    #region 140
                    case "140":

                        if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction == true &&
                            pVal.ItemUID == "47")
                        {
                            csVariablesGlobales.SboApp.MessageBox("La condición de pago ha cambiado. Revise la lista de precios.", 1, "Ok", "", "");
                        }

                        /*Creo el boton*/
                        if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            oItem = oForm.Items.Add("btnCarRef", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem.Left = oForm.Items.Item("14").Left + oForm.Items.Item("14").Width + 10;
                            oItem.Width = oForm.Items.Item("1").Width + 20;
                            oItem.Top = oForm.Items.Item("14").Top;
                            oItem.FontSize = oForm.Items.Item("1").FontSize;
                            oItem.Height = oForm.Items.Item("1").Height;
                            oButton = (SAPbouiCOM.Button)oForm.Items.Item("btnCarRef").Specific;
                            oButton.Caption = "Carga Referencias";
                            oItem.Visible = true;
                        }

                        if (pVal.EventType == BoEventTypes.et_CLICK && pVal.BeforeAction)
                        {
                            oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                            if (pVal.ItemUID == "btnCarRef")
                            {
                                string cClaveBase, cTipoBase, cReferencia;
                                oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("38").Specific);
                                try
                                {        /**/
                                    
                                    oForm.Freeze(true);
                                    for (int ncont = 1; ncont <= oMatrix.VisualRowCount; ncont++)
                                    {
                                        cTipoBase = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("43", ncont)).Value;
                                        if (cTipoBase == "17")
                                        {
                                            cClaveBase = ((SAPbouiCOM.EditText)oMatrix.GetCellSpecific("45", ncont)).Value;
                                            cReferencia = csUtilidades.DameValor("ORDR", "NumAtCard", "DocEntry = " + cClaveBase);
                                            if (cReferencia != "")
                                            {
                                                ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("163").Cells.Item(ncont).Specific)).String = cReferencia;
                                            }
                                        }
                                    }
                                    
                                }
                                catch (Exception ex)
                                {
                                    oForm.Freeze(false);
                                    csVariablesGlobales.SboApp.MessageBox("No puede modificar la referencia", 0, "", "", "");
                                }
                                oForm.Freeze(false);
                                                                
                            }
                        }
                        
                        break;
                    #endregion
                    #region 179
                    case "179":
                        if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction == true &&
                            pVal.ItemUID == "47")
                        {
                            csVariablesGlobales.SboApp.MessageBox("La condición de pago ha cambiado. Revise el descuento comercial.", 1, "Ok", "", "");
                        }
                        break;
                    #endregion
                }
                #region Descuentos
                if (csVariablesGlobales.CrearDtosEnDocumentos != "N")
                {
                    switch (Formulario)
                    {
                        #region "133", "149", "139", "140", "180", "133", "179", "142", "143", "141", "181", "65308", "65300", "60090", "60091", "65309", "65301", "60092"
                        case "133":
                        case "149":
                        case "139":
                        case "140":
                        case "180":
                        case "179":
                        case "142":
                        case "143":
                        case "141":
                        case "181":
                        case "65308":
                        case "65300":
                        case "60090":
                        case "60091":
                        case "65309":
                        case "65301":
                        case "60092": //Documentos
                            #region Descuentos
                            if (pVal.ColUID == "14")
                            {
                                if (pVal.ItemUID == "38" & pVal.EventType == BoEventTypes.et_VALIDATE & pVal.ItemChanged & !pVal.BeforeAction)
                                {
                                    if ((SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_ADD_MODE | (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                    {
                                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("14").Cells.Item(pVal.Row).Specific);
                                        string Precio = oEditText.String;
                                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_PRET").Cells.Item(pVal.Row).Specific);
                                        oEditText.String = csUtilidades.MonedaADouble(Precio).ToString();
                                        oForm.Freeze(true);
                                        if (csVariablesGlobales.RecalculoDePrecio)
                                        {
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO1").Cells.Item(pVal.Row).Specific);
                                            if (oEditText.Value != "0.0")
                                            {
                                                oEditText.String = "0";
                                            }
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO2").Cells.Item(pVal.Row).Specific);
                                            if (oEditText.Value != "0.0")
                                            {
                                                oEditText.String = "0";
                                            }
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO3").Cells.Item(pVal.Row).Specific);
                                            if (oEditText.Value != "0.0")
                                            {
                                                oEditText.String = "0";
                                            }
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO4").Cells.Item(pVal.Row).Specific);
                                            if (oEditText.Value != "0.0")
                                            {
                                                oEditText.String = "0";
                                            }
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_DTO5").Cells.Item(pVal.Row).Specific);
                                            if (oEditText.Value != "0.0")
                                            {
                                                oEditText.String = "0";
                                            }
                                            csVariablesGlobales.RecalculoDePrecio = true;
                                        }
                                        else
                                        {
                                            csVariablesGlobales.RecalculoDePrecio = true;
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("15").Cells.Item(pVal.Row).Specific);
                                            oEditText.String = "0";
                                        }
                                        oForm.Freeze(false);
                                    }
                                }
                            }
                            if (pVal.ColUID == "U_DOR_DTO1" | pVal.ColUID == "U_DOR_DTO2" | pVal.ColUID == "U_DOR_DTO3" | pVal.ColUID == "U_DOR_DTO4" | pVal.ColUID == "U_DOR_DTO5")
                            {
                                if (pVal.ItemUID == "38" & pVal.EventType == BoEventTypes.et_VALIDATE & pVal.ItemChanged & !pVal.BeforeAction)
                                {
                                    if ((SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_ADD_MODE | (SAPbouiCOM.BoFormMode)(pVal.FormMode) == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                    {
                                        oForm = csVariablesGlobales.SboApp.Forms.Item(FormUID);
                                        string Precio;
                                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific);
                                        if (Convert.ToDouble(oEditText.String) >= 0 && Convert.ToDouble(oEditText.String) <= 100)
                                        {
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_PRET").Cells.Item(pVal.Row).Specific);
                                            Precio = oEditText.String;
                                            oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific);
                                            if (csUtilidades.DameValor("OITM", "ItemCode", "ItemCode='" + oEditText.Value + "'") == "")
                                            {
                                                return;
                                            }
                                            switch (pVal.ColUID)
                                            {
                                                case "U_DOR_DTO1":
                                                case "U_DOR_DTO2":
                                                case "U_DOR_DTO3":
                                                case "U_DOR_DTO4":
                                                case "U_DOR_DTO5":
                                                    oForm.Freeze(true);
                                                    if (csVariablesGlobales.RecalculoDePrecio)
                                                    {
                                                        csVariablesGlobales.RecalculoDePrecio = false;
                                                        csUtilidades.CargarDescuentos(oMatrix, pVal, out Dto1, out Dto2, out Dto3, out Dto4, out Dto5, out dblPrecio);
                                                        csUtilidades.CalcularDescuentosVentas(dblPrecio, Dto1, Dto2, Dto3, Dto4, Dto5, oMatrix, pVal, csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_PARAM]", "U_Valor", "U_Concepto = 'Aplicar Dtos. En: Precio/Descuento'")); //"Descuento" o "Precio"
                                                        oForm.Freeze(false);
                                                        oEditText = (SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_DOR_PRET").Cells.Item(pVal.Row).Specific);
                                                        oEditText.String = csUtilidades.MonedaADouble(Precio).ToString();
                                                        csVariablesGlobales.RecalculoDePrecio = true;
                                                    }
                                                    else
                                                    {
                                                        csVariablesGlobales.RecalculoDePrecio = true;
                                                    }
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            csVariablesGlobales.SboApp.StatusBar.SetText("El descuento no puede ser: " + oEditText.String + ". Debe ser un valor entre 0 y 100", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
                                            oEditText.String = "";
                                        }
                                    }
                                }
                            }
                            #endregion
                            break;
                        #endregion
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "Ok", "", "");
                BubbleEvent = true;
            }
        }

        private void SBOApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                string Formulario = csUtilidades.DameValor("[@" + csVariablesGlobales.Prefijo + "_FORMS]", "U_FormSIA", "U_FormSAP='" + pVal.FormTypeEx + "'");
                switch (Formulario)
                {
                    #region 133, 179, 60090,  60091, 65300
                    case "133":
                    case "179":
                    case "60090":
                    case "60091":
                    case "65300":
                        if (pVal.EventType == BoEventTypes.et_FORM_DATA_ADD &&
                            pVal.BeforeAction && !pVal.ActionSuccess && pVal.Type != "112")
                        {
                            SAPbouiCOM.ComboBox oComboBox = null;
                            oForm = csVariablesGlobales.SboApp.Forms.Item(pVal.FormUID);
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                            string IndicadorImpuesto;// = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("18").Cells.
                                                     //  Item(1).Specific).Selected.Value;
                            string EsUE = "N";// = csUtilidades.DameValor("OVTG", "IsEC", "Code = '" + IndicadorImpuesto + "'");
                            string EsEX = "N";
                            for (int i = 1; i <= oMatrix.VisualRowCount - 1; i++)
                            {
                                IndicadorImpuesto = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("18").Cells.
                                                    Item(i).Specific).Selected.Value;
                                if ("EX" == IndicadorImpuesto)
                                {
                                    EsEX = "Y";
                                }
                                if ("Y" == csUtilidades.DameValor("OVTG", "IsEC", "Code = '" + IndicadorImpuesto + "'"))
                                {
                                    EsUE = "Y";
                                }
                                if (EsEX == EsUE && EsEX == "Y")
                                {
                                    csVariablesGlobales.SboApp.MessageBox("Indicadores de impuestos EX y UE mezclados", 1, "Ok", "", "");
                                    BubbleEvent = false;
                                    return;
                                }
                                //switch (EsUE)
                                //{
                                //    case "Y":
                                //        if ("EX" == ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("18").Cells.
                                //                    Item(i).Specific).Selected.Value)
                                //        {
                                //            csVariablesGlobales.SboApp.MessageBox("Indicadores de impuestos EX y UE mezclados", 1, "Ok", "", "");
                                //            BubbleEvent = false;
                                //            return;
                                //        }
                                //        break;
                                //    case "N":
                                //        IndicadorImpuesto = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("18").Cells.
                                //                            Item(i).Specific).Selected.Value;
                                //        if ("Y" == csUtilidades.DameValor("OVTG", "IsEC", "Code = '" + IndicadorImpuesto + "'") &&
                                //            IndicadorImpuesto == "EX")
                                //        {
                                //            csVariablesGlobales.SboApp.MessageBox("Indicadores de impuestos EX y UE mezclados", 1, "Ok", "", "");
                                //            BubbleEvent = false;
                                //            return;
                                //        }
                                //        break;
                                //}
                            }

                        }
                        break;
                    #endregion
                    case "60051":
                        //if (pVal.EventType == BoEventTypes.et_FORM_DATA_ADD &&
                        //    !pVal.BeforeAction && pVal.ActionSuccess)
                        //{
                        //    string Valor = ((SAPbouiCOM.ComboBox)oForm.Items.Item("4").Specific).Selected.Value;
                        //    if (Valor == "D" || Valor == "P")
                        //    {
                        //        SAPbobsCOM.JournalEntries oJournalEntries;
                        //        oJournalEntries = (SAPbobsCOM.JournalEntries)(csVariablesGlobales.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries));
                        //        //csVariablesGlobales.SboApp.MessageBox(csVariablesGlobales.oCompany.GetNewObjectKey(), 1, "Ok", "", "");
                        //        //MessageBox.Show(csVariablesGlobales.oCompany.GetNewObjectKey());
                        //        //oJournalEntries.GetByKey(Convert.ToInt32(csVariablesGlobales.oCompany.GetNewObjectKey()));
                        //    }
                        //}
                        break;
                    default:

                        break;
                }
            }
            catch (Exception ex)
            {
                if (ex.Message != "Form - Invalid Form")
                {
                    csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "Ok", "", "");
                }
                BubbleEvent = true;
            }
        }
    }
}
