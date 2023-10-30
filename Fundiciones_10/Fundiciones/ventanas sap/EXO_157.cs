using System;
using System.Linq;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_157
    {
        public EXO_157(){ }

        public bool ItemEvent(ItemEvent infoEvento)
        {
            switch (infoEvento.EventType)
            {
                case BoEventTypes.et_VALIDATE:
                    #region validate
                    if(!infoEvento.BeforeAction && infoEvento.ActionSuccess && !infoEvento.InnerEvent  && infoEvento.ItemChanged )
                    {
                        #region validate especifico para col precio
                        if (infoEvento.ItemUID == "3" && infoEvento.ColUID == "5" && infoEvento.ItemChanged)
                        {
                            Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                            ComboBox box = ((Matrix)oForm.Items.Item("3").Specific).GetCellSpecific("U_EXO_BLOQPRECIOEXO", infoEvento.Row);
                            box.Select("Y");
                        }
                        #endregion

                    }
                    #endregion
                    break;

                case BoEventTypes.et_FORM_LOAD:
                    #region visible
                    if (!infoEvento.BeforeAction)
                    {
                        Form oForm = Matriz.gen.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
                        Column col = ((Matrix)oForm.Items.Item("3").Specific).Columns.Item("U_EXO_BLOQPRECIOEXO");
                        col.ExpandType = BoExpandType.et_DescriptionOnly;
                        col.DisplayDesc = true;
                    }
                    #endregion
                    break;
                    
            }
            return true;
        }
    }
}