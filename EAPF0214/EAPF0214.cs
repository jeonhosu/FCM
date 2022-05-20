using System;
using System.Windows.Forms;

using Syncfusion.Windows.Forms;
using Syncfusion.Windows.Forms.Tools;
using Syncfusion.Windows.Forms.Grid;
using InfoSummit.Win.ControlAdv;

namespace EAPF0214
{
    public partial class EAPF0214 : Office2007Form
    {
        #region ----- Variables -----



        #endregion;

        #region ----- Constructor -----

        public EAPF0214()
        {
            InitializeComponent();
        }

        public EAPF0214(Form pMainForm, ISAppInterface pAppInterface)
        {
            InitializeComponent();

            this.MdiParent = pMainForm;

            isAppInterfaceAdv1.AppInterface = pAppInterface;
        }

        #endregion;

        #region ----- Private Methods ----

        private void SearchFromDataAdapter()
        {

            IDA_MASTER_NO_RULE.SetSelectParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            IDA_MASTER_NO_RULE.SetSelectParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);


            IDA_MASTER_NO_RULE.Fill();
        }

        #endregion;

        #region ----- Events -----

        private void EAPF0306_Load(object sender, EventArgs e)
        {

            ILD_MASTER_NO_RULE.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ILD_MASTER_NO_RULE.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);

            ILD_MASTER_TYPE.SetLookupParamValue("W_SOB_ID", isAppInterfaceAdv1.SOB_ID);
            ILD_MASTER_TYPE.SetLookupParamValue("W_ORG_ID", isAppInterfaceAdv1.ORG_ID);

            IDA_MASTER_NO_RULE.FillSchema();
        }

        private void isAppInterfaceAdv1_AppMainButtonClick(ISAppButtonEvents e)
        {
            if (this.IsActive)
            {
                if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Search)
                {
                    SearchFromDataAdapter();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddOver)
                {
                    if (IDA_MASTER_NO_RULE.IsFocused)
                    {
                        IDA_MASTER_NO_RULE.AddOver();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.AddUnder)
                {
                    if (IDA_MASTER_NO_RULE.IsFocused)
                    {
                        IDA_MASTER_NO_RULE.AddUnder();
                    }
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Update)
                {

                    object vUserId = isAppInterfaceAdv1.AppInterface.UserId;

                    IDA_MASTER_NO_RULE.SetInsertParamValue("P_SOB_ID", isAppInterfaceAdv1.SOB_ID);
                    IDA_MASTER_NO_RULE.SetInsertParamValue("P_ORG_ID", isAppInterfaceAdv1.ORG_ID);
                    IDA_MASTER_NO_RULE.SetInsertParamValue("P_USER_ID", isAppInterfaceAdv1.AppInterface.UserId);
                    IDA_MASTER_NO_RULE.Update();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Cancel)
                {
                    IDA_MASTER_NO_RULE.Cancel();
                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Delete)
                {

                }
                else if (e.AppMainButtonType == InfoSummit.Win.ControlAdv.ISUtil.Enum.AppMainButtonType.Print)
                {
                    //GridView(isGridAdvEx2, isDataAdapter2);
                }
            }
        }

        private void GridView(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, InfoSummit.Win.ControlAdv.ISDataAdapter pAdapter)
        {
            if (pAdapter != null)
            {
                pGrid.RowCount = pAdapter.OraSelectData.Rows.Count;
                pGrid.ColCount = pAdapter.OraSelectData.Columns.Count;

                foreach (System.Data.DataRow vRow in pAdapter.OraSelectData.Rows)
                {
                    foreach (System.Data.DataColumn vCol in pAdapter.OraSelectData.Columns)
                    {
                        int vRowIndex = vRow.Table.Rows.IndexOf(vRow);
                        int vColIndex = vRow.Table.Columns.IndexOf(vCol);

                        pGrid.SetCellValue(vRowIndex, vColIndex, vRow[vCol]);
                    }
                }
            }
        }

        #endregion;

    }
}