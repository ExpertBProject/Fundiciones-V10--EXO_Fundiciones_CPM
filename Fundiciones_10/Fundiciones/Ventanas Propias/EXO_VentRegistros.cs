using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    class EXO_VENREG
    {

        public EXO_VENREG()
        { }

        public EXO_VENREG(List<Matriz.Mensajes> ListaMensajes, string Titulo)
        {
           
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.Matrix oMatrix = null;
                #region CargoScreen

                SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.gen.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));
                string strXML = Utilidades.LeoFichEmbebido("Formularios.EXO_VenReg.srf");                
                oParametrosCreacion.XmlData = strXML;               
                try
                {
                    oForm = Matriz.gen.SBOApp.Forms.AddEx(oParametrosCreacion);
                }
                catch (Exception ex)
                {
                    Matriz.gen.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                }

                #endregion
                oForm.Title = Titulo;

                SAPbouiCOM.DataTable oTabla = oForm.DataSources.DataTables.Item("TablaReg");
                foreach (Matriz.Mensajes AuxMensaje in ListaMensajes)
                {
                    oTabla.Rows.Add();
                    oTabla.SetValue("Clave", oTabla.Rows.Count - 1, AuxMensaje.Clave);
                    oTabla.SetValue("Mensaje", oTabla.Rows.Count - 1, AuxMensaje.Mensaje.PadRight(254).Substring(0, 254).Trim());
                    oTabla.SetValue("Objeto", oTabla.Rows.Count - 1, AuxMensaje.Objeto );
                }


                #region binding
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLin").Specific;
                oMatrix.Columns.Item("V_1").DataBind.Bind("TablaReg", "Clave");
                oMatrix.Columns.Item("V_0").DataBind.Bind("TablaReg", "Mensaje");
                oMatrix.Columns.Item("V_2").DataBind.Bind("TablaReg", "Objeto");
                ((SAPbouiCOM.LinkedButton)oMatrix.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Order;
                #endregion

                oMatrix.Columns.Item("V_2").Visible = false;
                oMatrix.LoadFromDataSource();            
        }

        public bool ItemEvent(ItemEvent infoEvento)
        {

            switch (infoEvento.EventType)
            {
                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    {
                        if (infoEvento.ItemUID == "matLin" && infoEvento.BeforeAction)
                        {
                            SAPbouiCOM.Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

                            //infoEvento
                            SAPbouiCOM.Matrix oMatLin = (SAPbouiCOM.Matrix) oForm.Items.Item("matLin").Specific;

                            string cObjeto = ((SAPbouiCOM.EditText)oMatLin.GetCellSpecific("V_2", infoEvento.Row)).Value;
                            switch (cObjeto)
                            {
                                //Factura
                                case "13":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Invoice;                                    
                                    break;
                                //Abono
                                case "14":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_InvoiceCreditMemo;
                                    break;
                                //Albaran de venta
                                case "15":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_DeliveryNotes;
                                    break;
                                //Pedido de venta
                                case "17":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Order;
                                    break;

                                //Albaran de compra
                                case "20":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_GoodsReceiptPO;
                                    break;

                                //Entrada de mercancias
                                case "59":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_GoodsReceipt;
                                    break;

                                //Salida de mercancias
                                case "60":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_GoodsIssue;
                                    break;
                                
   
                                    //Asiento
                                case "30":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_JournalPosting;
                                    break;
                                //Documentos preliminares
                                case "112":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Drafts;
                                    break;
                                //Traslados
                                case "67":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_StockTransfers;
                                    break;
                                //Cobros
                                case "24":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Receipt;
                                    break;
                                //Cobros
                                case "4":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Items;
                                    break;                                                             
                                //Albaran inicial
                                case "IA":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_GoodsReceiptPO;
                                    break;                                
                                //entrada inicial
                                case "IE":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_GoodsReceipt;
                                    break;
                                default:

                                    return false;                                
                            }                            
                        }                        
                    }
                    break;

            }
            return true;
        }
    }
}
