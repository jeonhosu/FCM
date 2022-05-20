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

namespace FCMF0102
{
    public partial class FCMF0102 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public FCMF0102(Form pMainForm, ISAppInterface pAppInterface)
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
                    if (idaCUST_PARTY.IsFocused == true)
                    {
                        idaCUST_PARTY.AddOver();
                        GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaCUST_PARTY.IsFocused == true)
                    {
                        idaCUST_PARTY.AddUnder();
                        GRID_DefaultValue();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (idaCUST_PARTY.IsFocused == true)
                    {
                        idaCUST_PARTY.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaCUST_PARTY.IsFocused == true)
                    {
                        idaCUST_PARTY.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaCUST_PARTY.IsFocused == true)
                    {
                        idaCUST_PARTY.Delete();
                    }
                }
            }
        }

        #endregion;

        #region -- Data Find --
        private void CUST_PARTY_INQUIRY()
        {
            idaCUST_PARTY.Fill();
        }

        #endregion


        #region -- Default Value Setting --
        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            isgCUST_PARTY.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            isgCUST_PARTY.SetCellValue("ENABLED_FLAG", "Y");
        }
        #endregion


        private void FCMF0102_Load(object sender, EventArgs e)
        {
            idaCUST_PARTY.FillSchema();
        }

        private void isgCUST_PARTY_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            if (e.ColIndex == 3)
            {
                string V_Check_Result = null;
                idcPARTY_DESC.SetCommandParamValue("P_CUST_PARTY_DESC", e.NewValue);
                idcPARTY_DESC.ExecuteNonQuery();

                V_Check_Result = idcPARTY_DESC.GetCommandParamValue("X_CHECK_RESULT").ToString();

                
                if (V_Check_Result == 'N'.ToString())
                {
                    MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_90003", "&&FIELD_NAME:=Customer Party"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }
            }
        }

        private void idaCUST_PARTY_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=Customer Party"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }
    }
}