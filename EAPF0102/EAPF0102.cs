using System;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using InfoSummit.Win.ControlAdv;

namespace EAPF0102
{
    public partial class EAPF0102 : Office2007Form
    {
        #region ----- Variables -----

        //private int mGridCurrentCellValidatingCount = 0;

        #endregion;

        #region ----- Constructor -----

        public EAPF0102()
        {
            InitializeComponent();
        }

        public EAPF0102(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchFromDataAdapter()
        {
            IDA_ORG.Fill();
        }

        private bool IsAbleToNewAdded(ISDataAdapter pDataAdapter)
        {
            bool IsAble = false;

            if (pDataAdapter.CurrentRow.RowState == System.Data.DataRowState.Added)
            {
                IsAble = true;
            }

            return IsAble;
        }

        private bool IsCodeValidation(char vChar)
        {
            bool IsValidation = false;

            char[] vAlphaArray = new char[37] { '_', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

            int vCountFalse = 0;
            int vLength = vAlphaArray.Length;
            for (int vCol = 0; vCol < vLength; vCol++)
            {
                if (vAlphaArray[vCol] == vChar)
                {
                    vCountFalse++;
                }
            }

            if (vCountFalse > 0)
            {
                IsValidation = true;
            }

            return IsValidation;
        }

        private bool IsCode(string vText)
        {
            bool vIsCode = true;
            int vCountFalse = 0;
            int vLength = vText.Length;
            char[] vChars = vText.ToCharArray();

            for (int vCol = 0; vCol < vLength; vCol++)
            {
                bool isCode = IsCodeValidation(vChars[vCol]);
                if (isCode != true)
                {
                    vCountFalse++;
                }
            }

            if (vCountFalse > 0)
            {
                vIsCode = false;
            }

            return vIsCode;
        }

        private void ValidatingDate(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            ISGridAdvEx vGrid = pSender as ISGridAdvEx;

            int vColIndex_CAMType = vGrid.GetColumnToIndex("ORG_CODE");
            if (e.ColIndex == vColIndex_CAMType)
            {
                if (e.NewValue != null)
                {
                    Type vType = e.NewValue.GetType();
                    bool isNull1 = vType == Type.GetType("System.DBNull") ? true : false;
                    if (isNull1 == false)
                    {
                        bool isConvert = e.NewValue is string;
                        if (isConvert == true)
                        {
                            string vCodeString = e.NewValue as string;
                            bool isNull2 = string.IsNullOrEmpty(vCodeString);
                            if (isNull2 == false)
                            {
                                //bool vIsCode = IsCode(vCodeString);
                                //if (vIsCode == false)
                                //{
                                //    e.Cancel = true;

                                //    string vMessageGet = isMessageAdapter1.ReturnText("EAPP_10012");
                                //    string vMessageString = string.Format("{0}", vMessageGet);
                                //    MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                //}

                                ISCommonUtil.ISFunction.ISCode isCode = new ISCommonUtil.ISFunction.ISCode();
                                bool vIsCode = isCode.ISCheckCode(vCodeString);
                                if (vIsCode != true)
                                {
                                    e.Cancel = true;

                                    string vMessageGet = isMessageAdapter1.ReturnText("EAPP_10012"); //대문자 영어와 숫자를 조합한 코드만 입력하세요!
                                    string vMessageString = string.Format("{0}", vMessageGet);
                                    MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                            }
                            else
                            {
                                e.Cancel = true;

                                string vMessageGet = isMessageAdapter1.ReturnText("EAPP_90002");
                                string vMessageString = string.Format("{0}", vMessageGet);
                                MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
        }

        #endregion;

        #region ----- Events -----

        private void EAPF0102_Load(object sender, EventArgs e)
        {
            IDA_ORG.FillSchema();
            ILD_TRX_PRICE.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ILD_TRX_PRICE.SetLookupParamValue("W_LOOKUP_TYPE", "FG_PRICE");
            ILD_INV_PRICE.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ILD_INV_PRICE.SetLookupParamValue("W_LOOKUP_TYPE", "FG_PRICE");
            ILD_MAT_PRICE.SetLookupParamValue("W_LOOKUP_MODULE", "EAPP");
            ILD_MAT_PRICE.SetLookupParamValue("W_LOOKUP_TYPE", "MAT_PRICE");
        }

        private void isGridAdvEx1_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            ValidatingDate(pSender, e);
        }

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchFromDataAdapter();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    IDA_ORG.AddOver();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_ORG.AddUnder();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    if (IDA_ORG.IsFocused == true)
                    {
                        IDA_ORG.Update();
                    }
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_ORG.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    bool IsAble = IsAbleToNewAdded(IDA_ORG);
                    if (IsAble == false)
                    {
                        string vMessageGet = isMessageAdapter1.ReturnText("EAPP_10013");
                        string vMessageString = string.Format("{0}", vMessageGet);
                        MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    IDA_ORG.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }

        #endregion;

        private void ILA_FG_FIFO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FIFO.SetLookupParamValue("W_LOOKUP_TYPE", "FIFO_DIVISION");
        }

        private void ILA_MAT_FIFO_PrePopupShow(object pSender, ISLookupPopupShowEventArgs e)
        {
            ILD_FIFO.SetLookupParamValue("W_LOOKUP_TYPE", "FIFO_DIVISION");
        }

        private void ILA_FI_TAX_CODE_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_FI_COMMON.SetLookupParamValue("W_GROUP_CODE", "TAX_CODE");
            ILD_FI_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

        private void ILA_FI_OPERATION_DIVISION_RefreshLookupData(object pSender, ISRefreshLookupDataEventArgs e)
        {
            ILD_FI_COMMON.SetLookupParamValue("W_GROUP_CODE", "OPERATION_DIVISION");
            ILD_FI_COMMON.SetLookupParamValue("W_ENABLED_YN", "Y");
        }

    }
}
