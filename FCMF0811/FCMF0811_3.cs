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

namespace FCMF0811
{
    public partial class FCMF0811_3 : Office2007Form
    {
        #region ----- Variables -----

        private ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        private ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        public object Get_Save_Flag
        {
            get
            {
                return V_SAVE_FLAG.EditValue;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public FCMF0811_3(Form pMainForm, ISAppInterface pAppInterface
                        , object pTAX_DESC, object pTAX_CODE
                        , object pVAT_PERIOD_DESC
                        , object pVAT_DATE_FR, object pVAT_DATE_TO
                        , object pNOT_DED_TYPE
                        , object pNOT_DED_DESC, object pNOT_DED_CODE)
        {
            InitializeComponent();
            //this.MdiParent = pMainForm;
            isAppInterfaceAdv1.AppInterface = pAppInterface;

            T.EditValue = pTAX_CODE;
            V_TAX_DESC.EditValue = pTAX_DESC;

            V_VAT_PERIOD_DESC.EditValue = pVAT_PERIOD_DESC;

            V_VAT_DATE_FR.EditValue = pVAT_DATE_FR;
            V_VAT_DATE_TO.EditValue = pVAT_DATE_TO;

            V_NOT_DED_TYPE.EditValue = pNOT_DED_TYPE;
            V_NOT_DED_DESC.EditValue = pNOT_DED_DESC;
            V_NOT_DED_CODE.EditValue = pNOT_DED_CODE;

            V_ADJUST_TYPE.EditValue = "3";
        }

        #endregion;

        #region ----- Private Methods ----

        private void SEARCH_DB()
        {
            object vObject1 = V_TAX_DESC.EditValue;
            object vObject3 = V_VAT_PERIOD_DESC.EditValue;
            if (iConv.ISNull(vObject1) == string.Empty || iConv.ISNull(vObject3) == string.Empty)
            {
                //사업장, 과세년도, 신고기간구분은 필수 입니다.
                MessageBoxAdv.Show(isMessageAdapter1.ReturnText("FCM_10366"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            IDA_ADJUST_3.Fill();
        }

        private void SetCommonParameter(object pGroup_Code, object pEnabled_YN)
        {
            //ildCOMMON.SetLookupParamValue("W_GROUP_CODE", pGroup_Code);
            //ildCOMMON.SetLookupParamValue("W_ENABLED_YN", pEnabled_YN);
        }

        private void Set_GRID_STATUS_ROW()
        {
            if (IGR_ADJUST_3.RowCount < 1)
            {
                return;
            }
            int vSTATUS = 0;                // INSERTABLE, UPDATABLE;

            int vROW = IGR_ADJUST_3.RowIndex;
            object vNO_DED_CODE = IGR_ADJUST_3.GetCellValue("NO_DED_CODE");
            int vIDX_GL_AMOUNT = IGR_ADJUST_3.GetColumnToIndex("GL_AMOUNT");
            int vIDX_VAT_AMOUNT = IGR_ADJUST_3.GetColumnToIndex("VAT_AMOUNT");

            if (iConv.ISNull(vNO_DED_CODE) == "990")
            {
                vSTATUS = 0;
            }
            else
            {
                vSTATUS = 1;
            }

            IGR_ADJUST_3.GridAdvExColElement[vIDX_GL_AMOUNT].Insertable = vSTATUS;
            IGR_ADJUST_3.GridAdvExColElement[vIDX_GL_AMOUNT].Updatable = vSTATUS;

            IGR_ADJUST_3.GridAdvExColElement[vIDX_VAT_AMOUNT].Insertable = vSTATUS;
            IGR_ADJUST_3.GridAdvExColElement[vIDX_VAT_AMOUNT].Updatable = vSTATUS;
        }

        #endregion;

        #region ----- Events -----

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    
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
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                    
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Export)
                {
                    
                }
            }
        }

        #endregion;

        #region ----- Form Event -----

        private void FCMF0811_3_Load(object sender, EventArgs e)
        {
            IDA_ADJUST_3.FillSchema();
        }
        
        private void FCMF0811_3_Shown(object sender, EventArgs e)
        {
            SEARCH_DB();
            V_SAVE_FLAG.EditValue = "NONE";
        }

        private void IGR_ADJUST_3_CurrentCellChanged(object pSender, ISGridAdvExChangedEventArgs e)
        {
            int vIDX_SUPPLY_AMT = IGR_ADJUST_3.GetColumnToIndex("SUPPLY_AMT");
            int vIDX_VAT_AMT = IGR_ADJUST_3.GetColumnToIndex("VAT_AMT");
            int vIDX_CAL_TYPE = IGR_ADJUST_3.GetColumnToIndex("CAL_TYPE");
            int vIDX_TAX_SUPPLY_AMT = IGR_ADJUST_3.GetColumnToIndex("TAX_SUPPLY_AMT");
            int vIDX_NON_TAX_SUPPLY_AMT = IGR_ADJUST_3.GetColumnToIndex("NON_TAX_SUPPLY_AMT");

            if (e.ColIndex == vIDX_SUPPLY_AMT || e.ColIndex == vIDX_VAT_AMT || e.ColIndex == vIDX_CAL_TYPE
                || e.ColIndex == vIDX_TAX_SUPPLY_AMT || e.ColIndex == vIDX_NON_TAX_SUPPLY_AMT)
            {
                decimal vSUPPLY_AMT = iConv.ISDecimaltoZero(IGR_ADJUST_3.GetCellValue("SUPPLY_AMT"), 0);
                decimal vVAT_AMT = iConv.ISDecimaltoZero(IGR_ADJUST_3.GetCellValue("VAT_AMT"), 0);
                decimal vTAX_SUPPLY_AMT = iConv.ISDecimaltoZero(IGR_ADJUST_3.GetCellValue("TAX_SUPPLY_AMT"), 0);
                decimal vNON_TAX_SUPPLY_AMT = iConv.ISDecimaltoZero(IGR_ADJUST_3.GetCellValue("NON_TAX_SUPPLY_AMT"), 0);

                //(9)과세,면세사업 공통매입액 공급가액//
                if (e.ColIndex == vIDX_SUPPLY_AMT)
                {
                    vSUPPLY_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);                    
                }
                //(10)세액//
                if (e.ColIndex == vIDX_VAT_AMT)
                {
                    vVAT_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);                    
                }
                //(11)총공급가액등//
                if (e.ColIndex == vIDX_TAX_SUPPLY_AMT)
                {
                    vTAX_SUPPLY_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);
                }
                //(12)면세공급가액등//
                if (e.ColIndex == vIDX_NON_TAX_SUPPLY_AMT)
                {
                    vNON_TAX_SUPPLY_AMT = iConv.ISDecimaltoZero(e.NewValue, 0);
                }

                //(A)면세비율(%)//
                decimal vNON_TAX_RATE = 0;
                if (vTAX_SUPPLY_AMT != 0)
                {
                    vNON_TAX_RATE = Math.Floor(10000 * ((vNON_TAX_SUPPLY_AMT / vTAX_SUPPLY_AMT) * 100)) / 10000;
                }
                //(13)불공제매입세액//
                decimal vNO_VAT_AMT = Math.Floor((vVAT_AMT * (vNON_TAX_RATE / 100)));
                //(안분후공급가//
                decimal vDIVISION = 0.1M;
                decimal vADJUST_SUPPLY_AMT = Math.Floor(vNO_VAT_AMT / vDIVISION);

                IGR_ADJUST_3.SetCellValue("NON_TAX_RATE", vNON_TAX_RATE);
                IGR_ADJUST_3.SetCellValue("NOT_VAT_AMT", vNO_VAT_AMT);
                IGR_ADJUST_3.SetCellValue("ADJUST_SUPPLY_AMT", vADJUST_SUPPLY_AMT);
            }            
        }

        private void BTN_ADD_UNDER_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ADJUST_3.AddUnder();

            IGR_ADJUST_3.Focus();
        }

        private void BTN_UPDATE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ADJUST_3.Update();
        }

        private void BTN_DELETE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ADJUST_3.Delete();
        }
        
        private void BTN_CANCEL_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            IDA_ADJUST_3.Cancel();
        }

        private void BTN_CLOSE_ButtonClick(object pSender, EventArgs pEventArgs)
        {
            this.Close();
        }

        #endregion

        #region ----- Lookup Event -----

        private void ilaTAX_CODE_0_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            SetCommonParameter("TAX_CODE", "Y");
        }

        private void ilaPOP_VAT_REPORT_MNG_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            
        }

        #endregion

        #region ----- Grid Event -----
        
        private void igrSUM_NO_DEDUCTION_CurrentCellValidated(object pSender, ISGridAdvExValidatedEventArgs e)
        {
            if (IGR_ADJUST_3.RowCount < 1)
            {
                return;
            }

            decimal vAMOUNT = 0;
            int vIDX_GL_AMOUNT = IGR_ADJUST_3.GetColumnToIndex("GL_AMOUNT");
            int vIDX_VAT_AMOUNT = IGR_ADJUST_3.GetColumnToIndex("VAT_AMOUNT");

            Decimal vGL_RATE = iConv.ISDecimaltoZero(10);
            Decimal vVAT_RATE = iConv.ISDecimaltoZero(0.1);

            if (e.ColIndex == vIDX_GL_AMOUNT)
            {
                if (iConv.ISDecimaltoZero(IGR_ADJUST_3.GetCellValue("VAT_AMOUNT"), 0) == 0)
                {
                    vAMOUNT = vVAT_RATE * iConv.ISDecimaltoZero(e.CellValue, 0);
                    IGR_ADJUST_3.SetCellValue("VAT_AMOUNT", vAMOUNT);
                }
            }
            else if (e.ColIndex == vIDX_VAT_AMOUNT)
            {
                if (iConv.ISDecimaltoZero(IGR_ADJUST_3.GetCellValue("GL_AMOUNT"), 0) == 0)
                {
                    vAMOUNT = vGL_RATE * iConv.ISDecimaltoZero(e.CellValue, 0);
                    IGR_ADJUST_3.SetCellValue("GL_AMOUNT", vAMOUNT);
                }
            }
        }

        private void igrZERO_TAX_SPEC_CurrentCellAcceptedChanges(object pSender, ISGridAdvExChangedEventArgs e)
        {
            InfoSummit.Win.ControlAdv.ISGridAdvEx vGrid = pSender as InfoSummit.Win.ControlAdv.ISGridAdvEx;

            int vIndexColunm = vGrid.GetColumnToIndex("PUBLISH_DATE");

            if (e.ColIndex == vIndexColunm)
            {
                object vObject = vGrid.GetCellValue("PUBLISH_DATE");
                vGrid.SetCellValue("SHIPPING_DATE", vObject);
            }
        }

        #endregion

        #region ----- Adapter Event -----

        private void IDA_ADJUST_3_UpdateCompleted(object pSender)
        {
            V_SAVE_FLAG.EditValue = "SAVE";

            this.Close();
        }

        #endregion


    }
}