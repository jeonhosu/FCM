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
using ISCommonUtil;

namespace CSTF0105
{
    public partial class CSTF0105 : Office2007Form
    {
        ISFunction.ISConvert mCommonUtil = new ISFunction.ISConvert(); 

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public CSTF0105(Form pMainForm, ISAppInterface pAppInterface)
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
                    IDA_FROM_CC.Fill();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                        //IDA_OP_DIST_RULE.AddOver();
                        //GRID_DefaultValue();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                        //IDA_OP_DIST_RULE.AddUnder();
                        //GRID_DefaultValue();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                        IDA_FROM_CC.Update();
                        Change_Allow();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (IDA_FROM_CC.IsFocused == true)
                    {
                        IDA_FROM_CC.Cancel();
                    }

                    if (IDA_TO_CC.IsFocused == true)
                    {
                        IDA_TO_CC.Cancel();
                    }

                    if (IDA_TO_OPERATION.IsFocused == true)
                    {
                        IDA_TO_OPERATION.Cancel();
                    }
                    Change_Allow();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                        //IDA_OP_DIST_RULE.Delete();
                }
            }
        }

        #endregion;



        private void CSTF0105_Load(object sender, EventArgs e)
        {
            IDA_FROM_CC.FillSchema();
            V_DIRECT_TYPE.EditValue = Convert.ToString("A");
        }

        #region -- Default Value Setting --
        private void GRID_DefaultValue()
        {
            idcLOCAL_DATE.ExecuteNonQuery();
            ISG_FROM_CC.SetCellValue("EFFECTIVE_DATE_FR", idcLOCAL_DATE.GetCommandParamValue("X_LOCAL_DATE"));
            ISG_FROM_CC.SetCellValue("ENABLED_FLAG", "Y");
        }
        #endregion




        private void BT_ACCOUNT_LOAD_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            if (IDA_FROM_CC.Refillable == true)
            {
                IDC_CC_LOAD.ExecuteNonQuery();

                string X_RESULT_STATUS = mCommonUtil.ISNull(IDC_CC_LOAD.GetCommandParamValue("X_RESULT_STATUS"));
                string X_RESULT_MSG = mCommonUtil.ISNull(IDC_CC_LOAD.GetCommandParamValue("X_RESULT_MSG"));

                if (!IDC_CC_LOAD.ExcuteError && X_RESULT_STATUS != string.Empty && X_RESULT_STATUS == "S")
                {
                    IDT_CC_LOAD.Commit();
                    MessageBoxAdv.Show(X_RESULT_MSG, "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    IDA_FROM_CC.Fill();
                }
                else
                {
                    IDT_CC_LOAD.RollBack();
                    MessageBoxAdv.Show(X_RESULT_MSG, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("CST_10015", null), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void RD_ALL_CheckChanged(object sender, EventArgs e)
        {
            if (RD_ALL.Checked == true)
            {
                V_DIRECT_TYPE.EditValue = RD_ALL.CheckedString.ToString();
                IDA_FROM_CC.Fill();
            }
        }

        private void RD_INDIRECT_CheckChanged(object sender, EventArgs e)
        {
            if (RD_INDIRECT.Checked == true)
            {
                V_DIRECT_TYPE.EditValue = RD_INDIRECT.CheckedString.ToString();
                IDA_FROM_CC.Fill();
            }
        }

        private void RD_DIRECT_CheckChanged(object sender, EventArgs e)
        {
            if (RD_DIRECT.Checked == true)
            {
                V_DIRECT_TYPE.EditValue = RD_DIRECT.CheckedString.ToString();
                IDA_FROM_CC.Fill();
            }
        }

        private void V_CC_ENABLED_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            IDA_FROM_CC.Fill();
        }

        private void V_OP_ENABLED_FLAG_CheckedChange(object pSender, ISCheckEventArgs e)
        {
            IDA_FROM_CC.Fill();
        }

        #region
        private void Change_Allow()
        {
            if (IDA_FROM_CC.Refillable == false || IDA_TO_CC.Refillable == false || IDA_TO_OPERATION.Refillable == false)
            {
                RD_ALL.Enabled = false;
                RD_DIRECT.Enabled = false;
                RD_INDIRECT.Enabled = false;
                V_OP_ENABLED_FLAG.Enabled = false;
                V_CC_ENABLED_FLAG.Enabled = false;
            }
            else
            {
                RD_ALL.Enabled = true;
                RD_DIRECT.Enabled = true;
                RD_INDIRECT.Enabled = true;
                V_OP_ENABLED_FLAG.Enabled = true;
                V_CC_ENABLED_FLAG.Enabled = true;
            }

            RD_ALL.Refresh();
            RD_DIRECT.Refresh();
            RD_INDIRECT.Refresh();
            V_OP_ENABLED_FLAG.Refresh();
            V_CC_ENABLED_FLAG.Refresh();
        }
        
        private void ISG_FROM_CC_CurrentCellEditingComplete(object pSender, ISGridAdvExCellEditingEventArgs e)
        {
            Change_Allow();
        }

        private void ISG_TO_CC_CurrentCellEditingComplete(object pSender, ISGridAdvExCellEditingEventArgs e)
        {
            Change_Allow();
        }

        private void ISG_TO_OPERATION_CurrentCellEditingComplete(object pSender, ISGridAdvExCellEditingEventArgs e)
        {
            Change_Allow();
        }

        #endregion


    }
}