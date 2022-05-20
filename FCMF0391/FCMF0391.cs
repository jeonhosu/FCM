using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;
using ISCommonUtil;

namespace FCMF0391
{
    public partial class FCMF0391 : Office2007Form
    {
        #region ----- Variables -----
        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private Z4Mplus vLabelPrint = null;

        #endregion;

        #region ----- Constructor -----

        public FCMF0391()
        {
            InitializeComponent();
        }

        public FCMF0391(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();
            this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void PrintLabel()
        {
            isAppInterfaceAdv1.OnAppMessage("Label Printing Start...");
            System.Windows.Forms.Application.DoEvents();

            try
            {
                int vIndexCheckBox = igrASSET_MASTER.GetColumnToIndex("SELECT_CHECK_YN");
                int vTotalRow = igrASSET_MASTER.RowCount;
                object vManageNo = null;        //관리번호
                object vAssetName = null;       //자산명
                object vItemSpec = null;        //규격
                object vManageF = null;         //관리자(정)
                object vManageS = null;         //관리자(부)
                object vAcquireDate = null;     //취득일자
                object vUseDept = null;         //사용부서     

                idaZebra.Fill();

                //int vCountRow = idaZebra.OraSelectData.Rows.Count;
                //if (vCountRow > 0)
                //{
                    Zebra zpl = new Zebra(isAppInterfaceAdv1, idaZebra.OraSelectData);
                    for (int nRow = 0; nRow < vTotalRow; nRow++)
                    {
                        if((string)igrASSET_MASTER.GetCellValue(nRow, vIndexCheckBox) == "Y")
                        {
                            igrASSET_MASTER.CurrentCellMoveTo(nRow, 0);
                            igrASSET_MASTER.Focus();
                            igrASSET_MASTER.CurrentCellActivate(nRow, 0);

                            vManageNo = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("ASSET_CODE"));                            
                            vAssetName = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("ASSET_DESC"));
                            vItemSpec = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("ITEM_SPEC"));
                            vManageF = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("FIRST_USER_NAME"));
                            vManageS = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("SECOND_USER_NAME"));
                            vAcquireDate = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("ACQUIRE_DATE"));
                            vUseDept = igrASSET_MASTER.GetCellValue(nRow, igrASSET_MASTER.GetColumnToIndex("USE_DEPT_NAME"));

                            zpl.Printing(vManageNo, vAssetName, vItemSpec, vManageF, vManageS, vAcquireDate, vUseDept);
                        }
                    }                 

                    zpl.Dispose();
                //}
            }
            catch(System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    if (iString.ISNull(ASSET_CATEGORY_ID_0.EditValue) == string.Empty)
                    {
                        MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10101"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        ASSET_CATEGORY_NAME_0.Focus();
                        return;
                    };

                    IDA_ASSET_MASTER.Fill();
                    igrASSET_MASTER.Focus();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                   }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_ASSET_MASTER.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    //==============================================================================
                    // 1. 사용자가 위치 설정을 할 수 있도록 만든 로직의 함수이다.
                    // 2. 영문, 숫자 출력 가능, 한글 출력 불가능
                    //==============================================================================
                    //PrintLabel();

                    //==============================================================================
                    // 1. 사용자가 위치 설정은 할 수 없으며 출력 위치는 하드 코딩으로 설정해야 한다.
                    // 2. 영문, 숫자, 한글 출력 가능
                    //==============================================================================
                    LabelPrinting();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    
                }
            }
        }

        #endregion;

        private void FCMF0391_Load(object sender, EventArgs e)
        {
            vLabelPrint = new Z4Mplus(isAppInterfaceAdv1.AppInterface, igrASSET_MASTER, printDialog1, printPreviewDialog1);
        }

        private void FCMF0391_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                if (vLabelPrint != null)
                {
                    vLabelPrint.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }
        
        // 전체선택 버튼
        private void btnSELECT_ALL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < igrASSET_MASTER.RowCount; i++)
            {
                igrASSET_MASTER.SetCellValue(i, igrASSET_MASTER.GetColumnToIndex("SELECT_CHECK_YN"), "Y");
            }

            igrASSET_MASTER.LastConfirmChanges();
            IDA_ASSET_MASTER.OraSelectData.AcceptChanges();
            IDA_ASSET_MASTER.Refillable = true;
        }

        // 취소 버튼
        private void btnCONFIRM_CANCEL_0_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            for (int i = 0; i < igrASSET_MASTER.RowCount; i++)
            {
                igrASSET_MASTER.SetCellValue(i, igrASSET_MASTER.GetColumnToIndex("SELECT_CHECK_YN"), "N");
            }
        }

        private void ilaASSET_CATEGORY_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_AST_CATEGORY_LEVEL", 1);
            ILD_ASSET_CATEGORY.SetLookupParamValue("W_ENABLED_YN", "N");
        }

        private void ilaASSET_STATUS_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommon_Lookup_Parameter("ASSET_STATUS", "Y");
        }

        private void SetCommon_Lookup_Parameter(string pGroup_Code, string pEnabled_YN)
        {
            ILD_COMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            ILD_COMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void ASSET_CATEGORY_NAME_0_KeyDown(object pSender, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Control)
            {
                switch (e.KeyCode)
                {
                    case Keys.O:
                        Zebra_Position vPrintSetup = new Zebra_Position(isAppInterfaceAdv1.AppInterface);
                        vPrintSetup.Show();

                        break;
                }
            }

        }
/*
// 'Ctrl+O'버튼을 누르면 사용자가 위치 설정을 할 수 있다.
        private void FCMF0391_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Control)
            {
                switch (e.KeyCode)
                {
                    case Keys.O:
                        Zebra_Position vPrintSetup = new Zebra_Position(isAppInterfaceAdv1.AppInterface);
                        vPrintSetup.Show();

                        break;
                }
            }
        }
*/

        #region ----- Z4Mplus Label Printing Method ----

        private void LabelPrinting()
        {
            try
            {
                vLabelPrint.PRINTING();
            }
            catch (System.Exception ex)
            {
                string vMessage = string.Format("{0} - {1}", vLabelPrint.ErrorMessage, ex.Message);
                isAppInterfaceAdv1.OnAppMessage(ex.Message);
            }
        }

        #endregion;

        private void ilaUSE_DEPT_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_DEPT.SetLookupParamValue("W_GROUP_CODE", "FLOOR");
            ILD_DEPT.SetLookupParamValue("W_ENABLED_FLAG_YN", "N");
        }

        private void FCMF0391_Shown(object sender, EventArgs e)
        {
            START_DATE.EditValue = iDate.ISMonth_1st(DateTime.Today);
            END_DATE.EditValue = DateTime.Today;
        }

        private void igrASSET_MASTER_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_CHECK_FLAG = igrASSET_MASTER.GetColumnToIndex("SELECT_CHECK_YN");
            if (e.ColIndex == vIDX_CHECK_FLAG)
            {
                igrASSET_MASTER.LastConfirmChanges();
                IDA_ASSET_MASTER.OraSelectData.AcceptChanges();
                IDA_ASSET_MASTER.Refillable = true;
            }

        }
    }
}