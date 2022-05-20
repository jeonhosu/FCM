using System;
using ISCommonUtil;

namespace FCMF0516 
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iConv = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // ��Ʈ�� ����.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "BILL_1";
        private string mSourceSheet2 = "";
        private string mSourceSheet3 = "Sheet2";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;    // ���ο� ������ üũ.

        // �μ�� ���ο� �հ�.
        private int mCopyLineSUM = 0;

        // �μ� 1���� �ִ� �μ�����.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 1;
        private int mCopy_EndRow = 1;
        private int m1stLastRow = 26;       //ù�� ���� �μ� ����.
        //private int m2ndLastRow = 32;     //ù��� ���� �μ� ����.

        private int mCurrentRow = 15;       //���� �μ�Ǵ� row ��ġ.
        private int mDefaultPageRow = 14;   // ������ ������ PageRow �⺻��.

        private int mDUP_Add_Row = 30;      //���忡 ������ ������ �μ��Ҷ� �μ��� row ������        

        ////���հ� : �Ǽ�, ��ȭ�ݾ�, ��ȭ�ݾ�.
        //private decimal mTOT_COUNT = 0;
        //private decimal mTOT_AMOUNT = 0;

        #endregion;

        #region ----- Property -----

        public string ErrorMessage
        {
            get
            {
                return mMessageError;
            }
        }

        public string OpenFileNameExcel
        {
            set
            {
                mXLOpenFileName = value;
            }
        }

        #endregion;

        #region ----- Constructor -----

        public XLPrinting(InfoSummit.Win.ControlAdv.ISAppInterface pAppInterface, InfoSummit.Win.ControlAdv.ISMessageAdapter pMessageAdapter)
        {
            mPrinting = new XL.XLPrint();
            mAppInterface = pAppInterface;
            mMessageAdapter = pMessageAdapter;
        }

        #endregion;

        #region ----- XL File Open -----

        public bool XLFileOpen()
        {
            bool IsOpen = false;

            try
            {
                IsOpen = mPrinting.XLOpenFile(mXLOpenFileName);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
            }

            return IsOpen;
        }

        #endregion;

        #region ----- Dispose -----

        public void Dispose()
        {
            mPrinting.XLOpenFileClose();
            mPrinting.XLClose();
        }

        #endregion;

        #region ----- MaxIncrement Methods ----

        private int MaxIncrement(string pPathBase, string pSaveFileName)
        {// ���ϸ� �ڿ� �Ϸù�ȣ ����.
            int vMaxNumber = 0;
            System.IO.DirectoryInfo vFolder = new System.IO.DirectoryInfo(pPathBase);
            string vPattern = string.Format("{0}*", pSaveFileName);
            System.IO.FileInfo[] vFiles = vFolder.GetFiles(vPattern);

            foreach (System.IO.FileInfo vFile in vFiles)
            {
                string vFileNameExt = vFile.Name;
                int vCutStart = vFileNameExt.LastIndexOf(".");
                string vFileName = vFileNameExt.Substring(0, vCutStart);

                int vCutRight = 3;
                int vSkip = vFileName.Length - vCutRight;
                string vTextNumber = vFileName.Substring(vSkip, vCutRight);
                int vNumber = int.Parse(vTextNumber);

                if (vNumber > vMaxNumber)
                {
                    vMaxNumber = vNumber;
                }
            }

            return vMaxNumber;
        }

        #endregion;

        #region ----- Array Set 1 ----

        private void SetArray1(int pTabIndex, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// �׸����� �÷��� ���� �÷��ε��� �� ����
            pGDColumn = new int[9];
            pXLColumn = new int[9];
            // �׸��� or �ƴ��� ��ġ.
            pGDColumn[0] = pGrid.GetColumnToIndex("BILL_NUM");
            pGDColumn[1] = pGrid.GetColumnToIndex("DUE_DATE");
            pGDColumn[2] = pGrid.GetColumnToIndex("ISSUE_DATE");
            pGDColumn[3] = pGrid.GetColumnToIndex("PERIOD_DAY");
            pGDColumn[4] = pGrid.GetColumnToIndex("BILL_STATUS_DESC");
            pGDColumn[5] = pGrid.GetColumnToIndex("BANK_NAME");
            pGDColumn[6] = pGrid.GetColumnToIndex("VENDOR_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("BILL_AMOUNT");
            pGDColumn[8] = pGrid.GetColumnToIndex("REMARK");

            if (pTabIndex == 1)
            {
                // ������ �μ��ؾ� �� ��ġ.
                pXLColumn[0] = 1;
                pXLColumn[1] = 4;
                pXLColumn[2] = 8;
                pXLColumn[3] = 12;
                pXLColumn[4] = 14;
                pXLColumn[5] = 18;
                pXLColumn[6] = 23;
                pXLColumn[7] = 31;
                pXLColumn[8] = 35;
            }
            else if (pTabIndex == 2)
            {
                // ������ �μ��ؾ� �� ��ġ.
                pXLColumn[0] = 1;
                pXLColumn[1] = 7;
                pXLColumn[2] = 11;
                pXLColumn[3] = 15;
                pXLColumn[4] = 17;
                pXLColumn[5] = 20;
                pXLColumn[6] = 24;
                pXLColumn[7] = 32;
                pXLColumn[8] = 36;
            }
        }

        #endregion;

        #region ----- Array Set 2 ----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// �׸����� �÷��� ���� �÷��ε��� �� ����
            pGDColumn = new int[9];
            pXLColumn = new int[9];
            // �׸��� or �ƴ��� ��ġ.
            pGDColumn[0] = pGrid.GetColumnToIndex("BILL_NUM");
            pGDColumn[1] = pGrid.GetColumnToIndex("DUE_DATE");
            pGDColumn[2] = pGrid.GetColumnToIndex("ISSUE_DATE");
            pGDColumn[3] = pGrid.GetColumnToIndex("PERIOD_DAY");
            pGDColumn[4] = pGrid.GetColumnToIndex("BILL_STATUS_DESC");
            pGDColumn[5] = pGrid.GetColumnToIndex("BANK_NAME");
            pGDColumn[6] = pGrid.GetColumnToIndex("VENDOR_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("BILL_AMOUNT");
            pGDColumn[8] = pGrid.GetColumnToIndex("REMARK");

            // ������ �μ��ؾ� �� ��ġ.
            pXLColumn[0] = 1;
            pXLColumn[1] = 7;
            pXLColumn[2] = 11;
            pXLColumn[3] = 15;
            pXLColumn[4] = 17;
            pXLColumn[5] = 20;
            pXLColumn[6] = 24;
            pXLColumn[7] = 32;
            pXLColumn[8] = 36;
        }

        #endregion;

        #region ----- Array Set 2  : Adapter ����� ----

        //private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{// �ƴ����� table ��.
        //    pGDColumn = new int[10];
        //    pXLColumn = new int[10];

        //    pGDColumn[0] = pTable.Columns.IndexOf("PO_TYPE_NAME");
        //    pGDColumn[1] = pTable.Columns.IndexOf("DISPLAY_NAME");
        //    pGDColumn[2] = pTable.Columns.IndexOf("PO_DATE");
        //    pGDColumn[3] = pTable.Columns.IndexOf("PO_NO");
        //    pGDColumn[4] = pTable.Columns.IndexOf("SUPPLIER_SHORT_NAME");
        //    pGDColumn[5] = pTable.Columns.IndexOf("PRICE_TERM_NAME");
        //    pGDColumn[6] = pTable.Columns.IndexOf("PAYMENT_METHOD_NAME");
        //    pGDColumn[7] = pTable.Columns.IndexOf("PAYMENT_TERM_NAME");
        //    pGDColumn[8] = pTable.Columns.IndexOf("REMARK");
        //    pGDColumn[9] = pTable.Columns.IndexOf("STEP_DESCRIPTION");


        //    pXLColumn[0] = 9;   //PO_TYPE_NAME
        //    pXLColumn[1] = 25;  //DISPLAY_NAME
        //    pXLColumn[2] = 42;  //PO_DATE
        //    pXLColumn[3] = 54;  //PO_NO
        //    pXLColumn[4] = 9;   //SUPPLIER_SHORT_NAME
        //    pXLColumn[5] = 35;  //PRICE_TERM_NAME
        //    pXLColumn[6] = 14;  //PAYMENT_METHOD_NAME
        //    pXLColumn[7] = 41;  //PAYMENT_TERM_NAME
        //    pXLColumn[8] = 7;   //REMARK
        //    pXLColumn[9] = 49;  //�ݾ�
        //}

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {// ���ڿ� ���� üũ �� �ش� �� ����.
            bool vIsConvert = false;
            pConvertString = string.Empty;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is string;
                    if (vIsConvert == true)
                    {
                        pConvertString = pObject as string;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertNumber(object pObject, out decimal pConvertDecimal)
        {// ���� ���� üũ �� �ش� �� ����.
            bool vIsConvert = false;
            pConvertDecimal = 0m;

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is decimal;
                    if (vIsConvert == true)
                    {
                        decimal vIsConvertNum = (decimal)pObject;
                        pConvertDecimal = vIsConvertNum;
                    }
                }

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        private bool IsConvertDate(object pObject, out System.DateTime pConvertDateTimeShort)
        {// ��¥ ���� üũ �� �ش� �� ����.
            bool vIsConvert = false;
            pConvertDateTimeShort = new System.DateTime();

            try
            {
                if (pObject != null)
                {
                    vIsConvert = pObject is System.DateTime;
                    if (vIsConvert == true)
                    {
                        System.DateTime vDateTime = (System.DateTime)pObject;
                        pConvertDateTimeShort = vDateTime;
                    }
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vIsConvert;
        }

        #endregion;



        #region ----- Excel Export Write Methods ----

        public int ExportWrite(string pTerritory, InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID)
        {// ���� ȣ��Ǵ� �κ�.
            string vMessage = string.Empty;
            string vVisible_YN = "0";
            
            int vCurrentCol = 1;
            int vTotalRow = pGRID.RowCount;
            int vTotalCol = pGRID.ColCount;
            int mPromptRow = 4;

            decimal vNumberValue = 0;

            object vDecimalDigit = 0;
            object vColumnType = null;
            object vValue = null;
            object vPrintValue = null;

            //�μ� ���� ����//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 10;
            mCopy_EndRow = 36;

            mCurrentRow = mPromptRow + 1;
                    
            try
            {
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.                

                if (vTotalRow > 0)
                {
                    #region ----- Write Page Copy(SourceSheet => TargetSheet) ----
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� �ٿ� �ִ´�.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet3, 1);

                    #endregion;

                    for (int c = 0; c < vTotalCol; c++)
                    {// ������Ʈ ǥ��.
                        vVisible_YN = iConv.ISNull(pGRID.GridAdvExColElement[c].Visible, "0");
                        if (vVisible_YN == "1")
                        {
                            if (pTerritory == "TL1_KR")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL1_KR;
                            }
                            else if (pTerritory == "TL2_CN")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL2_CN;
                            }
                            else if (pTerritory == "TL3_VN")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL3_VN;
                            }
                            else if (pTerritory == "TL4_JP")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL4_JP;
                            }
                            else if (pTerritory == "TL5_XAA")
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].TL5_XAA;
                            }

                            if (iConv.ISNull(vValue) == string.Empty)
                            {
                                vValue = pGRID.GridAdvExColElement[c].HeaderElement[0].Default;
                            }
                            vCurrentCol = vCurrentCol + 1;
                            mPrinting.XLSetCell(mPromptRow, vCurrentCol, vValue);
                        }
                    }

                    mPrinting.XLCellAlignmentHorizontal(mPromptRow, 1, mPromptRow, vCurrentCol, "C");
                    mCopy_EndCol = vCurrentCol;  // copy ���� ����.
                    vCurrentCol = 1;
                    for (int r = 0; r < vTotalRow; r++)
                    {//Row
                        for (int c = 0; c < vTotalCol; c++)
                        {//Col
                            vVisible_YN = iConv.ISNull(pGRID.GridAdvExColElement[c].Visible, "0");
                            if (vVisible_YN == "1")
                            {
                                vCurrentCol = vCurrentCol + 1;
                                vValue = pGRID.GetCellValue(r, c);
                                vColumnType = pGRID.GridAdvExColElement[c].ColumnType;
                                vDecimalDigit = pGRID.GridAdvExColElement[c].DecimalDigits;
                                if (iConv.ISNull(vColumnType) == "NumberEdit")
                                {
                                    try
                                    {
                                        vNumberValue = iConv.ISDecimaltoZero(vValue);
                                        if (iConv.ISNumtoZero(vDecimalDigit) > 0)
                                        {
                                            vPrintValue = string.Format("{0:###,###,###,###,###,###,###,###,##0.####}", vNumberValue);
                                        }
                                        else
                                        {
                                            vPrintValue = string.Format("{0:###,###,###,###,###,###,###,###,##0}", vNumberValue);
                                        }
                                    }
                                    catch
                                    {
                                        vPrintValue = vValue;
                                    }
                                    mPrinting.XLCellAlignmentHorizontal(mCurrentRow, vCurrentCol, mCurrentRow, vCurrentCol, "R");
                                }
                                else
                                {
                                    vPrintValue = vValue;
                                }
                                mPrinting.XLSetCell(mCurrentRow, vCurrentCol, vPrintValue);
                            }
                            vMessage = String.Format("Writing - [{0}/{1}]", r, vTotalRow);
                            mAppInterface.OnAppMessageEvent(vMessage);
                            System.Windows.Forms.Application.DoEvents();
                        }
                        vCurrentCol = 1;
                        mCurrentRow = mCurrentRow + 1;
                    }
                    mPrinting.XLColumnAutoFit(1, 1, mCurrentRow, mCopy_EndCol);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- Header Export Write Method ----

        public void Header_ExportWrite(Object pW_PERIOD_FR, Object pW_PERIOD_TO)
        {// ��� �μ�.
            int vXLine = 0;
            int vXLColumn = 0;
            String PEROID = String.Format("{0}~{1}", pW_PERIOD_FR, pW_PERIOD_TO);
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet3);

                //��¥.
                vXLine = 3;
                vXLColumn = 2;
                mPrinting.XLSetCell(vXLine, vXLColumn, PEROID);


            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;


        #region ----- Excel Wirte MAIN Methods ----

        //�׸��� 
        public int WriteMain(int pTabIndex
                            , string pDUE_DATE_FR
                            , string pDUE_DATE_TO
                            , DateTime pPrint_Datetime 
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// ���� ȣ��Ǵ� �κ�.
            string vMessage = string.Empty;

            int vTotalRow = 0;              //�� Row ��
            int vPageRowCount = 0;          //Page�� �μ�Ǵ� Row �� 
            //decimal vTotalPageNumber = 0;   //�� �μ� Page ��(�ø� �Լ� parameter Type)        
            
            //string vPrint_Value = string.Empty;
            //string vWorkCenter_Code = string.Empty;     //�۾��� �ڵ� 
            decimal vPO_REQ_AMOUNT = 0;     //���ſ�û ���� �ݾ� �հ� 

            int[] vGDColumn;
            int[] vXLColumn;

            // ���� SHEET ���� //
            if (pTabIndex == 1)
            {
                mSourceSheet1 = "BILL_1";
            }
            else if (pTabIndex == 2)
            {
                mSourceSheet1 = "BILL_2";
            }

            //�μ� ���� ����//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 48;
            mCopy_EndRow = 36;

            m1stLastRow = 35;                   //ù�� ���� �μ� ����.
            vPageRowCount = 4;                  // ù�忡 ���ؼ��� ����row���� üũ.

            mDefaultPageRow = 4;               // ������ ������ PageCount �⺻��.
            mCurrentRow = 5;                   // ���� �μ�Ǵ� row ��ġ.             

            // �μ⹰ ���� : ��� - ���� ������ ��� //
            try
            {
                XLHeader(pDUE_DATE_FR, pDUE_DATE_TO, pPrint_Datetime);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            // �μ� : ���� //
            try
            {
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.
                //�� row�� 
                vTotalRow = pGrid.RowCount;
                //vTotalPageNumber = Math.Ceiling((iConv.ISDecimaltoZero(vTotalRow * 2) / iConv.ISDecimaltoZero(m1stLastRow)));  //�μ� row�� 2�� �����ϹǷ�..

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                
                //---- Line Write Start ----
                if (vTotalRow > 0)
                {
                    //�׸��� �� ���� col index �迭 ����
                    SetArray1(pTabIndex, pGrid, out vGDColumn, out vXLColumn);

                    // ������ �����ؼ� Ÿ�꽬Ʈ�� ����/�ٿ� �ֱ�//
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);
                    mPrinting.XLActiveSheet(mTargetSheet);

                    for(int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow + 1, vTotalRow);

                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        //�۾��庰�� �μ⸦ ���� �ҽ�(�ٸ� �۾����� ��� �ٸ� ������ �μ�)   
                        //if (vWorkCenter_Code != string.Empty && vWorkCenter_Code != iConv.ISNull(vRow_L["WORKCENTER_CODE"]))
                        //{
                        //    //�ٸ� �۾��� : ���ο� ������ üũ �� ����.
                        //    mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + 1;
                        //    vPageRowCount = m1stLastRow;

                        //    IsNewPage(vPageRowCount);
                        //    if (mIsNewPage == true)
                        //    {
                        //        mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + mDefaultPageRow - 1;    // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                        //        vPageRowCount = mDefaultPageRow;
                        //    }
                        //}

                        //// �۾��� �� ���ڵ� �μ� //
                        //if (vWorkCenter_Code == string.Empty || vWorkCenter_Code != iConv.ISNull(vRow_L["WORKCENTER_CODE"]))
                        //{
                        //    //��û���� 
                        //    vPrint_Value = string.Format("{0}", vRow_L["REQ_ORDER_NO"]);
                        //    mPrinting.XLSetCell((mCurrentRow - 3), 27, vPrint_Value);

                        //    //���ó 
                        //    vPrint_Value = string.Format("{0}", vRow_L["WORKCENTER_DESCRIPTION"]);
                        //    mPrinting.XLSetCell((mCurrentRow - 2), 6, vPrint_Value);

                        //    //���ó �ڵ�
                        //    if (iConv.ISNull(vRow_L["WORKCENTER_CODE"]) == string.Empty)
                        //    {
                        //        vPrint_Value = string.Empty;
                        //    }
                        //    else
                        //    {
                        //        vPrint_Value = string.Format("*{0}*", vRow_L["WORKCENTER_CODE"]);
                        //    }
                        //    mPrinting.XLSetCell((mCurrentRow - 2), 13, vPrint_Value);
                        //}
                        
                        mCurrentRow = XLLine(vRow, pGrid, vGDColumn, vXLColumn);         // ���� �μ��ϴ� �Լ� 
                        mCurrentRow = mCurrentRow + 1;              // �μ� ���� ���� 
                        vPageRowCount = vPageRowCount + 1;          // �μ� ���� ī��Ʈ 

                        //vWorkCenter_Code = iConv.ISNull(vRow_L["WORKCENTER_CODE"]);  //�۾����ڵ� 

                        if (vRow == vTotalRow)
                        {
                            //// ������ ������ �̸� ó���� ���� ���     
                            ////�հ� �μ� 
                            //mCurrentRow = mCurrentRow + ((m1stLastRow - vPageRowCount) * 2);

                            //mPrinting.XLSetCell(mCurrentRow, 22, string.Format("{0:###,###}", vPO_REQ_AMOUNT));
                            //mPrinting.XLSetCell(mCurrentRow + 30, 22, string.Format("{0:###,###}", vPO_REQ_AMOUNT));  //add 
                        }
                        else
                        {
                            // ���ο� ������ üũ �� ����.
                            IsNewPage(mPageNumber, vPageRowCount);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;    // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = 4;
                            }
                        }
                    }
                    // Page ǥ�� //
                    XLPageNum_Total(mTargetSheet, mPageNumber);
                }
                //---- Line Write End ----
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        //Adapter �̿� 
        public int WriteMain(InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_HEADER
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_LINE)
        {// ���� ȣ��Ǵ� �κ�.
            string vMessage = string.Empty;

            int vRow = 0;                   //�μ�Ǵ� Data Row ��ġ 
            int vTotalRow = 0;              //�� Row ��
            int vPageRowCount = 0;          //Page�� �μ�Ǵ� Row �� 
            decimal vTotalPageNumber = 0;   //�� �μ� Page ��(�ø� �Լ� parameter Type)           

            string vPrint_Value = string.Empty;
            string vWorkCenter_Code = string.Empty;     //�۾��� �ڵ� 

            //�μ� ���� ����//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 45;
            mCopy_EndRow = 34;
            m1stLastRow = 33;                   //ù�� ���� �μ� ����.

            mDefaultPageRow = 6;                // ������ ������ PageCount �⺻��.
            mCurrentRow = 6;                    // ���� �μ�Ǵ� row ��ġ.
            vPageRowCount = mCurrentRow;    // ù�忡 ���ؼ��� ����row���� üũ.

            // �μ⹰ ���� : ��� - ���� ������ ��� //
            try
            {
                //�μ⹰ ��� : ���� ��Ʈ�� �����Ͱ��� ���� 
                foreach (System.Data.DataRow vRow1 in pIDA_HEADER.CurrentRows)
                {
                    XLHeader(vRow1);
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            // �μ� : ���� //
            try
            {
                //�� row�� 
                vTotalRow = pIDA_LINE.CurrentRows.Count;
                vTotalPageNumber = Math.Ceiling((iConv.ISDecimaltoZero(vTotalRow) / iConv.ISDecimaltoZero(m1stLastRow)));

                //mPageTotalNumber = vTotal1ROW / vBy;  // ���� �μ� ��� / �� ��� ǥ�� ����.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? ���� �տ� �� �����̰� : �������� ���� ��, �ڰ� ����.

                //---- Line Write Start ----
                if (vTotalRow > 0)
                {
                    // ������ �����ؼ� Ÿ�꽬Ʈ�� ����/�ٿ� �ֱ�//
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);
                    mPrinting.XLActiveSheet(mTargetSheet);

                    foreach (System.Data.DataRow vRow2 in pIDA_LINE.CurrentRows)
                    {
                        vRow = vRow + 1;
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);

                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        if (vWorkCenter_Code != string.Empty && vWorkCenter_Code != iConv.ISNull(vRow2["WORKCENTER_CODE"]))
                        {
                            //�ٸ� �۾��� : ���ο� ������ üũ �� ����.
                            mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + 1;
                            vPageRowCount = m1stLastRow;

                            IsNewPage(mPageNumber, vPageRowCount);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + mDefaultPageRow - 1;    // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }

                        // �۾��� �� ���ڵ� �μ� //
                        if (vWorkCenter_Code == string.Empty || vWorkCenter_Code != iConv.ISNull(vRow2["WORKCENTER_CODE"]))
                        {
                            //��û���� 
                            vPrint_Value = string.Format("{0}", vRow2["REQ_ORDER_NO"]);
                            mPrinting.XLSetCell((mCurrentRow - 3), 27, vPrint_Value);

                            //���ó 
                            vPrint_Value = string.Format("{0}", vRow2["WORKCENTER_DESCRIPTION"]);
                            mPrinting.XLSetCell((mCurrentRow - 2), 6, vPrint_Value);

                            //���ó �ڵ�
                            if (iConv.ISNull(vRow2["WORKCENTER_CODE"]) == string.Empty)
                            {
                                vPrint_Value = string.Empty;
                            }
                            else
                            {
                                vPrint_Value = string.Format("*{0}*", vRow2["WORKCENTER_CODE"]);
                            }
                            mPrinting.XLSetCell((mCurrentRow - 2), 13, vPrint_Value);
                        }

                        mCurrentRow = mCurrentRow + 1;
                        vPageRowCount = vPageRowCount + 1;

                        mCurrentRow = XLLine(vRow2);        // ���� �μ��ϴ� �Լ� 

                        vWorkCenter_Code = iConv.ISNull(vRow2["WORKCENTER_CODE"]);  //�۾����ڵ� 

                        if (vRow == vTotalRow)
                        {
                            // ������ ������ �̸� ó���� ���� ���                                                       
                        }
                        else
                        {
                            // ���ο� ������ üũ �� ����.
                            IsNewPage(mPageNumber, vPageRowCount);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;    // ������ �μ�� �ش� �������� ���۵Ǵ� ��ġ.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }

                    //// Page ǥ�� //
                    //for (int R = 1; R <= vTotalPageNumber; R++)
                    //{
                    //    mPrinting.XLSetCell(((m1stLastRow + 1) * R), 22, string.Format("[ {0} / {1} ]", R, vTotalPageNumber));
                    //}
                }
                //---- Line Write End ----
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }

        #endregion;

        #region ----- HEADER/LINE Excel Write -----

        #region ----- Header1 Write Method ----

        private void XLHeader(string pDUE_DATE_FR
                            , string pDUE_DATE_TO
                            , DateTime pPrint_Datetime
                            )
        {// ��� �μ�.
            int vXLine = 0;
            int vXLColumn = 0;

            object vObject = null;
            string vPrintValue = string.Empty;
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // ���� ���� �Ⱓ
                vObject = string.Format("{0} ~ {1}", pDUE_DATE_FR, pDUE_DATE_TO);
                vPrintValue = iConv.ISNull(vObject);
                vXLine = 3;
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue); 

                //�μ� �Ͻ�
                vPrintValue = string.Format("[{0:yyyy-MM-dd HH:mm:ss}]", pPrint_Datetime);
                vXLine = 36;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue); 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        //���޹��� ������ �̿��� �μ� 
        private void XLHeader(System.Data.DataRow pRow)
        {// ��� �μ�.
            int vXLine = 0;
            int vXLColumn = 0;

            object vObject = null;
            string vPrintValue = string.Empty;
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // PICKING NO
                vObject = pRow["PICKING_ORDER_NO"];
                vPrintValue = iConv.ISNull(vObject);
                vXLine = 3;
                vXLColumn = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue);

                // barcode 
                if (vPrintValue == string.Empty)
                {
                    vPrintValue = string.Empty;
                }
                else
                {
                    vPrintValue = string.Format("*{0}*", vPrintValue);
                }
                vXLine = 1;
                vXLColumn = 32;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue);

                // PICKING DATE
                vObject = pRow["PICKING_DATE"];
                vPrintValue = iConv.ISNull(vObject);
                vXLine = 5;
                vXLColumn = 6;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue);

                //����Ͻ� 
                vObject = pRow["PRINT_DATETIME"];
                vPrintValue = iConv.ISNull(vObject);
                vXLine = 34;
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Excel Write [Line] Method -----

        private int XLLine(int pRow, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {
            //pXLine : ������ �μ��ؾ� �ϴ� ��. pGDColumn : �׸��� ��ġ, pXLColumn : ���� ��ġ.

            int vXLine = mCurrentRow; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 1;

            // �μ� ���� //
            object vObject = null;
            string vPrint_Value = string.Empty;

            try
            { // Ÿ�ٽ�Ʈ Active.
                mPrinting.XLActiveSheet(mTargetSheet);

                //BILL_NUM 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[0]);
                vPrint_Value = string.Format("{0}", vObject);                
                vXLColumn = pXLColumn[0];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //DUE DATE 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[1]);
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = pXLColumn[1];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //ISSUE_DATE                 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[2]);
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = pXLColumn[2];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //PERIOD_DAY 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[3]);
                if (iConv.ISDecimaltoZero(vObject, 0) == 0)
                {
                    vPrint_Value = string.Empty;
                }
                else
                {
                    vPrint_Value = string.Format("{0:###,###,###,###,###,###.#####}", vObject);
                }
                vXLColumn = pXLColumn[3];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //BILL_STATUS_DESC 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[4]);
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = pXLColumn[4];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //BANK_NAME                 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[5]);
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = pXLColumn[5];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //VENDOR_NAME                 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[6]);
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = pXLColumn[6];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //BILL_AMOUNT                 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[7]);
                if (iConv.ISDecimaltoZero(vObject, 0) == 0)
                {
                    vPrint_Value = string.Empty;
                }
                else
                {
                    vPrint_Value = string.Format("{0:###,###,###,###,###,###.#####}", vObject);
                }
                vXLColumn = pXLColumn[7];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //REMARK                 
                vObject = pGrid.GetCellValue(pRow, pGDColumn[8]);
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = pXLColumn[8];
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
            return vXLine;
        }

        private int XLLine(System.Data.DataRow pRow)
        {
            //pXLine : ������ �μ��ؾ� �ϴ� ��. pGDColumn : �׸��� ��ġ, pXLColumn : ���� ��ġ.

            int vXLine = mCurrentRow; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLColumn = 1;

            // �μ� ���� //
            object vObject = null;
            string vPrint_Value = string.Empty;

            try
            { // Ÿ�ٽ�Ʈ Active.
                mPrinting.XLActiveSheet(mTargetSheet);

                //ITEM CODE 
                vObject = pRow["PRT_ITEM_CODE"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //�����                 
                vObject = pRow["PRT_ITEM_DESCRIPTION"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //����԰�                 
                vObject = pRow["PRT_ITEM_SPECIFICATION"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 12;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //UOM                   
                vObject = pRow["REQ_UOM_CODE"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //��û �ܷ� 
                vObject = pRow["PRT_REMAIN_QTY"];
                if (iConv.ISDecimaltoZero(vObject, 0) == 0)
                {
                    vPrint_Value = string.Empty;
                }
                else
                {
                    vPrint_Value = string.Format("{0:###,###,###,###,###,###.#####}", vObject);
                }
                vXLColumn = 23;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //���� 
                vObject = pRow["BOX_SEQ_NUM"];
                if (iConv.ISDecimaltoZero(vObject, 0) == 0)
                {
                    vPrint_Value = string.Empty;
                }
                else
                {
                    vPrint_Value = string.Format("{0:###,###}", vObject);
                }
                vXLColumn = 26;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //�ڽ���ȣ                 
                vObject = pRow["ONHAND_BOX_NO"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);


                //���                 
                vObject = pRow["BOX_ONHAND_QTY"];
                if (iConv.ISDecimaltoZero(vObject, 0) == 0)
                {
                    vPrint_Value = string.Empty;
                }
                else
                {
                    vPrint_Value = string.Format("{0:###,###,###,###,###,###.#####}", vObject);
                }
                vXLColumn = 34;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //PICKING ��                 
                vObject = pRow["BOX_PICKING_QTY"];
                if (iConv.ISDecimaltoZero(vObject, 0) == 0)
                {
                    vPrint_Value = string.Empty;
                }
                else
                {
                    vPrint_Value = string.Format("{0:###,###,###,###,###,###.#####}", vObject);
                }
                vXLColumn = 37;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            return vXLine;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        //private int XLTOTAL_Line(int pXLine)
        //{// pGridRow : �׸����� ���� �д� ��, pXLine : ������ �μ��ؾ� �ϴ� ��. pGDColumn : �׸��� ��ġ, pXLColumn : ���� ��ġ.
        //    int vXLine = pXLine; //������ ������ ǥ�õǴ� �� ��ȣ
        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // ������ �����ؼ� Ÿ�� �� ������ ����.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //�հ�
        //        vXLColumnIndex = 2;
        //        mPrinting.XLCellMerge(pXLine, vXLColumnIndex, pXLine, 35, true);
        //        vConvertString = string.Format("{0}", "��");
        //        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

        //        //17 - �Ǽ�
        //        vXLColumnIndex = 36;
        //        IsConvert = IsConvertNumber(mTOT_COUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }

        //        //31 - ���ް���
        //        vXLColumnIndex = 42;
        //        IsConvert = IsConvertNumber(mTOT_AMOUNT, out vConvertDecimal);
        //        if (IsConvert == true)
        //        {
        //            vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        else
        //        {
        //            vConvertString = string.Empty;
        //            mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
        //        }
        //        //-------------------------------------------------------------------
        //        vXLine = vXLine + 1;
        //        //-------------------------------------------------------------------
        //    }
        //    catch (System.Exception ex)
        //    {
        //        mMessageError = ex.Message;
        //        mAppInterface.OnAppMessageEvent(mMessageError);
        //        System.Windows.Forms.Application.DoEvents();
        //    }

        //    pXLine = vXLine;

        //    return pXLine;
        //}

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNum_Seq(string pActiveSheet, int pPageNumber)
        {// ���������� ������Ʈ �����ϱ� ���� ������Ʈ�� ����ϰ� ��Ʈ�� �����Ѵ�.

            int vXLRow = 33; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLCol = 46;

            try
            { // ������ �����ؼ� Ÿ�� �� ������ ����.(
                mPrinting.XLActiveSheet(pActiveSheet);
                mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        private void XLPageNum_Total(string pActiveSheet, int pTotalPageNum)
        {// ���������� ������Ʈ �����ϱ� ���� ������Ʈ�� ����ϰ� ��Ʈ�� �����Ѵ�.

            int vXLRow = 36; //������ ������ ǥ�õǴ� �� ��ȣ
            int vXLCol = 22;
            
            try
            { // ������ �����ؼ� Ÿ�� �� ������ ����.(
                mPrinting.XLActiveSheet(pActiveSheet);

                for (int R = 1; R <= pTotalPageNum; R++)
                {
                    mPrinting.XLSetCell(vXLRow, vXLCol, string.Format("[ {0} / {1} ]", R, pTotalPageNum));
                    vXLRow = vXLRow + mCopy_EndRow;
                }
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageNumber, int pPageRowCount)
        {
            if (pPageRowCount == m1stLastRow)
            { // pPrintingLine : ���� ��µ� ��.                
                mIsNewPage = true;

                mCurrentRow = mCopy_EndRow * pPageNumber;
                mCurrentRow = mCurrentRow + 1;

                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //������ ActiveSheet�� ������ ����  ������ ����
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            // page�� ǥ��.
            mPageNumber = mPageNumber + 1;

            //[����], [Sheet2.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(pActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[���], [Sheet1.Cell("A1:AS67")], ���� ��Ʈ���� ���� ������ ���ȣ, 
            //���� ��Ʈ���� ���� ������ ����ȣ, ���� ��Ʈ���� ���� ������ ���ȣ, ���� ��Ʈ���� ���� ������ ����ȣ
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // ����.

            return vPasteEndRow;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(string pPrint_Type, int pPageSTART, int pPageEND, int pPrintCopies)
        {
            if (pPrint_Type == "PREVIEW")
            {
                mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, pPrintCopies);
                mAppInterface.OnAppMessageEvent("Ready for printing [Preview]");
            }
            else
            {
                mPrinting.XLPrinting(pPageSTART, pPageEND, pPrintCopies);
                mAppInterface.OnAppMessageEvent("Ready for printing [Printer]");
            }
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            mAppInterface.OnAppMessageEvent("Ready for printing [Excel File]");
            //System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            //int vMaxNumber = MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
