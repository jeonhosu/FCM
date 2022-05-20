using System;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using InfoSummit.Win.ControlAdv;


namespace EAPF0101
{
    public partial class EAPF0101 : Office2007Form
    {
        #region ----- Variables -----

        private int mGridCurrentCellValidatingCount = 0;

        #endregion;

        #region ----- Constructor -----

        public EAPF0101()
        {
            InitializeComponent();
        }

        public EAPF0101(Form pMainForm, ISAppInterface pAppInterface)
        {
            //[2010-07-23]
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchFromDataAdapter()
        {
            IDA_SET_OF_BOOKS.Fill();
        }

        private int GetIndexISGridAdvExColElement(ISGenericCollection<ISGridAdvExColElement> pGridAdvExColElement, string pColumnType, string pDataColumn)
        {
            int vIndex = 0;
            string vColumnType = string.Empty;
            string vDataColumn = string.Empty;

            foreach (ISGridAdvExColElement ce in pGridAdvExColElement)
            {
                vColumnType = ce.ColumnType.ToString();
                vDataColumn = ce.DataColumn.ToString();


                if (vColumnType == pColumnType && vDataColumn == pDataColumn)
                {
                    vIndex = int.Parse(ce.Ordinal.ToString());
                }
            }

            return vIndex;
        }

        private bool IsNumAlpha(string vText)
        {
            bool vIsNumAlpha = false;
            int vCountNum = 0;
            int vCountChar = 0;
            int vLenght = vText.Length;
            char[] vChars = vText.ToCharArray();

            for (int vCol = 0; vCol < vLenght; vCol++)
            {
                bool isAlpha = char.IsUpper(vChars[vCol]);
                if (isAlpha == true)
                {
                    vCountChar++;
                }
                bool isNumber = char.IsNumber(vChars[vCol]);
                if (isNumber == true)
                {
                    vCountNum++;
                }
            }

            if (vCountChar > 0 && vCountNum > 0)
            {
                vIsNumAlpha = true;
            }

            return vIsNumAlpha;
        }

        private void GridCurrentCellValidating(ISGridAdvExValidatingEventArgs e)
        {
            string vColumnType = "TextEdit";
            string vDataColumn = "SOB_CODE";
            int vIndexColumn = GetIndexISGridAdvExColElement(IGR_SET_OF_BOOKS.GridAdvExColElement, vColumnType, vDataColumn);

            if (e.ColIndex == vIndexColumn)
            {
                if (e.NewValue != null)
                {
                    string vMessageString = string.Empty;
                    string vText = e.NewValue.ToString();
                    bool vIsNull = string.IsNullOrEmpty(vText);
                    if (vIsNull == true)
                    {
                        e.Cancel = true;

                        vMessageString = string.Format("{0}", isMessageAdapter1.ReturnText("EAPP_90002"));
                        MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        bool vIsNumAlpha = IsNumAlpha(vText);
                        if (vIsNumAlpha == false)
                        {
                            e.Cancel = true;

                            if (mGridCurrentCellValidatingCount == 1)
                            {
                                vMessageString = string.Format("[{0}]{1}", vText, isMessageAdapter1.ReturnText("EAPP_10012"));
                                MessageBoxAdv.Show(vMessageString, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                mGridCurrentCellValidatingCount = 0;
                            }
                            else
                            {
                                mGridCurrentCellValidatingCount++;
                            }
                        }
                    }
                }
            }
        }

        #endregion;

        #region ----- Events -----

        private void EAPF0101_Load(object sender, EventArgs e)
        {
            IDA_SET_OF_BOOKS.FillSchema();
        }

        private void isGridAdvEx1_CurrentCellValidating(object pSender, ISGridAdvExValidatingEventArgs e)
        {
            GridCurrentCellValidating(e);
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
                    IDA_SET_OF_BOOKS.AddOver();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    IDA_SET_OF_BOOKS.AddUnder();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Update)
                {
                    IDA_SET_OF_BOOKS.Update(); 
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_SET_OF_BOOKS.Cancel();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Delete)
                {
                    IDA_SET_OF_BOOKS.Delete();
                }
                else if (e.AppMainButtonType == ISUtil.Enum.AppMainButtonType.Print)
                {
                }
            }
        }

        #endregion;

        private void IDA_SET_OF_BOOKS_UpdateCompleted(object pSender)
        {
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10010"), isMessageAdapter1.ReturnText("EAPP_10206"));
        }

        private void IDA_SET_OF_BOOKS_PreDelete(ISPreDeleteEventArgs e)
        {
            MessageBoxAdv.Show(isMessageAdapter1.ReturnText("EAPP_10208"), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            e.Cancel = true;
        }
    }
}
