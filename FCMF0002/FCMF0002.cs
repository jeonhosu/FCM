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

namespace FCMF0002
{
    public partial class FCMF0002 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public FCMF0002(Form pMainForm, ISAppInterface pAppInterface)
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
                    if (idaSUPPLIER_TYPE.IsFocused == true)
                    {
                        idaSUPPLIER_TYPE.AddOver();
                        GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaSUPPLIER_TYPE.IsFocused == true)
                    {
                        idaSUPPLIER_TYPE.AddUnder();
                        GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaSUPPLIER_TYPE.IsFocused == true)
                    {
                        idaSUPPLIER_TYPE.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaSUPPLIER_TYPE.IsFocused == true)
                    {
                        idaSUPPLIER_TYPE.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaSUPPLIER_TYPE.IsFocused == true)
                    {
                        idaSUPPLIER_TYPE.Delete();
                    }
                }
            }
        }

        #endregion;

        #region -- Data Find --
        private void CUST_PARTY_INQUIRY()
        {
            idaSUPPLIER_TYPE.Fill();
        }

        #endregion

        private void FCMF0002_Load(object sender, EventArgs e)
        {
            idaSUPPLIER_TYPE.FillSchema();
        }

        #region -- Default Value Setting --
        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            isgSUPPLIER_TYPE.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            isgSUPPLIER_TYPE.SetCellValue("ENABLED_FLAG", "Y");
        }
        #endregion

        private void isgSUPPLIER_TYPE_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == 5)
            {
                string V_Check_Result = null;
                idcSUPPLIER_CODE.SetCommandParamValue("P_SUPPLIER_TYPE_CODE", e.NewValue);
                idcSUPPLIER_CODE.ExecuteNonQuery();

                V_Check_Result = idcSUPPLIER_CODE.GetCommandParamValue("X_CHECK_RESULT").ToString();


                if (V_Check_Result == 'N'.ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90003", "&&FIELD_NAME:=Supplier Type Code"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
        }


        private void idaSUPPLIER_TYPE_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Supplier Type"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
    }
}