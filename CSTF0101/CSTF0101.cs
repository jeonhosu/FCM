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

namespace CSTF0101
{
    public partial class CSTF0101 : Office2007Form
    {
        ISCommonUtil.ISFunction.ISConvert iConvert = new ISFunction.ISConvert();

        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public CSTF0101(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void idaFACTORY_ExcuteQuery()
        {
            idaFACTORY.Fill();
        }

        private void Factory_Setting()
        {
            iedCOST_FACTORY_CODE.LookupAdapter = null;
        }

        private void Division_Setting()
        {
            isGridAdvEx1.SetCellValue("COST_DIVISION_CODE",iedCOST_FACTORY_CODE.EditValue);
        }

        private void Class_Setting()
        {
            isGridAdvEx2.SetCellValue("COST_OP_CLASS_CODE", isGridAdvEx1.GetCellValue("COST_DIVISION_CODE"));
        }

        private void Section_Setting()
        {
            isGridAdvEx3.SetCellValue("COST_OP_SECTION_CODE", isGridAdvEx2.GetCellValue("COST_OP_CLASS_CODE"));
        }

        private void Center_Setting()
        {
            isGridAdvEx4.SetCellValue("COST_CENTER_CODE", isGridAdvEx3.GetCellValue("COST_OP_SECTION_CODE"));
            isGridAdvEx4.SetCellValue("MFG_TYPE_LCODE", isGridAdvEx1.GetCellValue("MFG_TYPE_LCODE"));
            isGridAdvEx4.SetCellValue("EFFECTIVE_DATE_FR", DateTime.Today);
            isGridAdvEx4.SetCellValue("ENABLED_FLAG", "Y");
        }

        #endregion;

        #region ----- Events -----

        private void CSTF0101_Load(object sender, EventArgs e)
        {
            idaFACTORY.FillSchema();
            idaDIVISION.FillSchema();
            idaCLASS.FillSchema();
            idaSECTION.FillSchema();
            idaCENTER.FillSchema();
        }

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    idaFACTORY.Fill();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (idaFACTORY.IsFocused == true)
                    {
                        idaFACTORY.AddOver();
                        Factory_Setting();
                    }
                    else if (idaDIVISION.IsFocused == true)
                    {
                        idaDIVISION.AddOver();
                        Division_Setting();
                    }
                    else if (idaCLASS.IsFocused == true)
                    {
                        idaCLASS.AddOver();
                        Class_Setting();
                    }
                    else if (idaSECTION.IsFocused == true)
                    {
                        idaSECTION.AddOver();
                        Section_Setting();
                    }
                    else if (idaCENTER.IsFocused == true)
                    {
                        idaCENTER.AddOver();
                        Center_Setting();
                    }

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (idaFACTORY.IsFocused == true)
                    {
                        idaFACTORY.AddUnder();
                        Factory_Setting();
                    }
                    else if (idaDIVISION.IsFocused == true)
                    {
                        idaDIVISION.AddUnder();
                        Division_Setting();
                    }
                    else if (idaCLASS.IsFocused == true)
                    {
                        idaCLASS.AddUnder();
                        Class_Setting();
                    }
                    else if (idaSECTION.IsFocused == true)
                    {
                        idaSECTION.AddUnder();
                        Section_Setting();
                    }
                    else if (idaCENTER.IsFocused == true)
                    {
                        idaCENTER.AddUnder();
                        Center_Setting();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    idaFACTORY.Update();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    if (idaFACTORY.IsFocused == true)
                    {
                        idaFACTORY.Cancel();
                    }
                    else if (idaDIVISION.IsFocused == true)
                    {
                        idaDIVISION.Cancel();
                    }
                    else if (idaCLASS.IsFocused == true)
                    {
                        idaCLASS.Cancel();
                    }
                    else if (idaSECTION.IsFocused == true)
                    {
                        idaSECTION.Cancel();
                    }
                    else if (idaCENTER.IsFocused == true)
                    {
                        idaCENTER.Cancel();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    if (idaFACTORY.IsFocused == true)
                    {
                        idaFACTORY.Delete();
                    }
                    else if (idaDIVISION.IsFocused == true)
                    {
                        idaDIVISION.Delete();
                    }
                    else if (idaCLASS.IsFocused == true)
                    {
                        idaCLASS.Delete();
                    }
                    else if (idaSECTION.IsFocused == true)
                    {
                        idaSECTION.Delete();
                    }
                    else if (idaCENTER.IsFocused == true)
                    {
                        idaCENTER.Delete();
                    }                          
                }
            }
        }

        private void idaFACTORY_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=" + isGroupBox1.PromptText.ToString()), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaDIVISION_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=" + isGroupBox2.PromptText.ToString()), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaCLASS_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=" + isGroupBox3.PromptText.ToString()), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaSECTION_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=" + isGroupBox4.PromptText.ToString()), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaCENTER_PreDelete(ISPreDeleteEventArgs e)
        {
            if (e.Row.RowState != DataRowState.Added)
            {
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10029", "&&VALUE:=" + isGroupBox5.PromptText.ToString()), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);        // 모듈 코드 입력
                e.Cancel = true;
                return;
            }
        }

        private void idaFACTORY_ExcuteKeySearch(object pSender)
        {
            idaFACTORY_ExcuteQuery();
        }

        private void isGridAdvEx1_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            switch (isGridAdvEx1.GridAdvExColElement[e.ColIndex].DataColumn.ToString())
            {
                case "MFG_TYPE_LCODE":
                    for (int vLoop = 0; vLoop < isGridAdvEx4.RowCount; vLoop++)
                    {
                        isGridAdvEx4.SetCellValue("MFG_TYPE_LCODE", isGridAdvEx1.GetCellValue("MFG_TYPE_LCODE"));
                    }

                    break;

                default:
                    break;
            }
        }

        private void iedCOST_FACTORY_CODE_PreKeySearch(object pSender)
        {
            iedCOST_FACTORY_CODE.LookupAdapter = ilaFACTORY;
        }

        private void iedCOST_FACTORY_CODE_PostKeySearch(object pSender)
        {
            iedCOST_FACTORY_CODE.LookupAdapter = null;
        }

        private void iedCOST_FACTORY_CODE_CancelKeySearch(object pSender)
        {
            iedCOST_FACTORY_CODE.LookupAdapter = null;
        }

        private void ILA_FI_OPERATION_DIVISION_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_FI_OPERATION_DIVISION.SetLookupParamValue("W_GROUP_CODE", "OPERATION_DIVISION");
            ILD_FI_OPERATION_DIVISION.SetLookupParamValue("W_ENABLED_YN", "Y");
        }
        
        #endregion;


    }
}