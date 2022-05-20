using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace FCMF0001
{
    public partial class FCMF0001 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public FCMF0001(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----



        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    CUST_PARTY_INQUIRY();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaSUPPLIER_CLASS.IsFocused == true)
                    {
                        idaSUPPLIER_CLASS.AddOver();
                        GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaSUPPLIER_CLASS.IsFocused == true)
                    {
                        idaSUPPLIER_CLASS.AddUnder();
                        GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaSUPPLIER_CLASS.IsFocused == true)
                    {
                        idaSUPPLIER_CLASS.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaSUPPLIER_CLASS.IsFocused == true)
                    {
                        idaSUPPLIER_CLASS.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaSUPPLIER_CLASS.IsFocused == true)
                    {
                        idaSUPPLIER_CLASS.Delete();
                    }
                }
            }
        }

        #endregion;

        #region -- Data Find --
        private void CUST_PARTY_INQUIRY()
        {
            idaSUPPLIER_CLASS.Fill();
        }

        #endregion

        private void FCMF0001_Load(object sender, EventArgs e)
        {
            idaSUPPLIER_CLASS.FillSchema();
        }


        #region -- Default Value Setting --
        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            isgSUPPLIER_CLASS.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            isgSUPPLIER_CLASS.SetCellValue("ENABLED_FLAG", "Y");
        }
        #endregion

        private void isgSUPPLIER_CLASS_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == 3)
            {
                string V_Check_Result = null;
                idcSUPPLIER_CLASS_CODE.SetCommandParamValue("P_SUPPLIER_CLASS_CODE", e.NewValue);
                idcSUPPLIER_CLASS_CODE.ExecuteNonQuery();

                V_Check_Result = idcSUPPLIER_CLASS_CODE.GetCommandParamValue("X_CHECK_RESULT").ToString();


                if (V_Check_Result == 'N'.ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90003", "&&FIELD_NAME:=Supplier Class Code"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
        }

        private void idaSUPPLIER_CLASS_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Supplier Class"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
    }
}