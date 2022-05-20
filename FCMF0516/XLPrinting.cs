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

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "BILL_1";
        private string mSourceSheet2 = "";
        private string mSourceSheet3 = "Sheet2";

        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;    // 새로운 페이지 체크.

        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 1;
        private int mCopy_EndRow = 1;
        private int m1stLastRow = 26;       //첫장 최종 인쇄 라인.
        //private int m2ndLastRow = 32;     //첫장외 최종 인쇄 라인.

        private int mCurrentRow = 15;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 14;   // 페이지 증가후 PageRow 기본값.

        private int mDUP_Add_Row = 30;      //한장에 동일한 내용을 인쇄할때 인쇄할 row 증가분        

        ////총합계 : 건수, 외화금액, 원화금액.
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
        {// 파일명 뒤에 일련번호 증가.
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
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[9];
            pXLColumn = new int[9];
            // 그리드 or 아답터 위치.
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
                // 엑셀에 인쇄해야 할 위치.
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
                // 엑셀에 인쇄해야 할 위치.
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
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[9];
            pXLColumn = new int[9];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("BILL_NUM");
            pGDColumn[1] = pGrid.GetColumnToIndex("DUE_DATE");
            pGDColumn[2] = pGrid.GetColumnToIndex("ISSUE_DATE");
            pGDColumn[3] = pGrid.GetColumnToIndex("PERIOD_DAY");
            pGDColumn[4] = pGrid.GetColumnToIndex("BILL_STATUS_DESC");
            pGDColumn[5] = pGrid.GetColumnToIndex("BANK_NAME");
            pGDColumn[6] = pGrid.GetColumnToIndex("VENDOR_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("BILL_AMOUNT");
            pGDColumn[8] = pGrid.GetColumnToIndex("REMARK");

            // 엑셀에 인쇄해야 할 위치.
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

        #region ----- Array Set 2  : Adapter 적용시 ----

        //private void SetArray2(System.Data.DataTable pTable, out int[] pGDColumn, out int[] pXLColumn)
        //{// 아답터의 table 값.
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
        //    pXLColumn[9] = 49;  //금액
        //}

        #endregion;

        #region ----- IsConvert Methods -----

        private bool IsConvertString(object pObject, out string pConvertString)
        {// 문자열 여부 체크 및 해당 값 리턴.
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
        {// 숫자 여부 체크 및 해당 값 리턴.
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
        {// 날짜 여부 체크 및 해당 값 리턴.
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
        {// 실제 호출되는 부분.
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

            //인쇄 범위 설정//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 10;
            mCopy_EndRow = 36;

            mCurrentRow = mPromptRow + 1;
                    
            try
            {
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.                

                if (vTotalRow > 0)
                {
                    #region ----- Write Page Copy(SourceSheet => TargetSheet) ----
                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet3, 1);

                    #endregion;

                    for (int c = 0; c < vTotalCol; c++)
                    {// 프롬프트 표시.
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
                    mCopy_EndCol = vCurrentCol;  // copy 영역 지정.
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
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;
            String PEROID = String.Format("{0}~{1}", pW_PERIOD_FR, pW_PERIOD_TO);
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet3);

                //날짜.
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

        //그리드 
        public int WriteMain(int pTabIndex
                            , string pDUE_DATE_FR
                            , string pDUE_DATE_TO
                            , DateTime pPrint_Datetime 
                            , InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;

            int vTotalRow = 0;              //총 Row 수
            int vPageRowCount = 0;          //Page당 인쇄되는 Row 수 
            //decimal vTotalPageNumber = 0;   //총 인쇄 Page 수(올림 함수 parameter Type)        
            
            //string vPrint_Value = string.Empty;
            //string vWorkCenter_Code = string.Empty;     //작업장 코드 
            decimal vPO_REQ_AMOUNT = 0;     //구매요청 예상 금액 합계 

            int[] vGDColumn;
            int[] vXLColumn;

            // 원본 SHEET 설정 //
            if (pTabIndex == 1)
            {
                mSourceSheet1 = "BILL_1";
            }
            else if (pTabIndex == 2)
            {
                mSourceSheet1 = "BILL_2";
            }

            //인쇄 범위 설정//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 48;
            mCopy_EndRow = 36;

            m1stLastRow = 35;                   //첫장 최종 인쇄 라인.
            vPageRowCount = 4;                  // 첫장에 대해서는 시작row부터 체크.

            mDefaultPageRow = 4;               // 페이지 증가후 PageCount 기본값.
            mCurrentRow = 5;                   // 현재 인쇄되는 row 위치.             

            // 인쇄물 종류 : 헤더 - 라인 관계일 경우 //
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

            // 인쇄 : 라인 //
            try
            {
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.
                //총 row수 
                vTotalRow = pGrid.RowCount;
                //vTotalPageNumber = Math.Ceiling((iConv.ISDecimaltoZero(vTotalRow * 2) / iConv.ISDecimaltoZero(m1stLastRow)));  //인쇄 row가 2씩 증가하므로..

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                
                //---- Line Write Start ----
                if (vTotalRow > 0)
                {
                    //그리드 및 엑셀 col index 배열 저장
                    SetArray1(pTabIndex, pGrid, out vGDColumn, out vXLColumn);

                    // 원본을 복사해서 타깃쉬트에 복사/붙여 넣기//
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);
                    mPrinting.XLActiveSheet(mTargetSheet);

                    for(int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow + 1, vTotalRow);

                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        //작업장별로 인쇄를 위한 소스(다른 작업장인 경우 다른 페이지 인쇄)   
                        //if (vWorkCenter_Code != string.Empty && vWorkCenter_Code != iConv.ISNull(vRow_L["WORKCENTER_CODE"]))
                        //{
                        //    //다른 작업장 : 새로운 페이지 체크 및 생성.
                        //    mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + 1;
                        //    vPageRowCount = m1stLastRow;

                        //    IsNewPage(vPageRowCount);
                        //    if (mIsNewPage == true)
                        //    {
                        //        mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + mDefaultPageRow - 1;    // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                        //        vPageRowCount = mDefaultPageRow;
                        //    }
                        //}

                        //// 작업장 및 바코드 인쇄 //
                        //if (vWorkCenter_Code == string.Empty || vWorkCenter_Code != iConv.ISNull(vRow_L["WORKCENTER_CODE"]))
                        //{
                        //    //요청정보 
                        //    vPrint_Value = string.Format("{0}", vRow_L["REQ_ORDER_NO"]);
                        //    mPrinting.XLSetCell((mCurrentRow - 3), 27, vPrint_Value);

                        //    //출고처 
                        //    vPrint_Value = string.Format("{0}", vRow_L["WORKCENTER_DESCRIPTION"]);
                        //    mPrinting.XLSetCell((mCurrentRow - 2), 6, vPrint_Value);

                        //    //출고처 코드
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
                        
                        mCurrentRow = XLLine(vRow, pGrid, vGDColumn, vXLColumn);         // 실제 인쇄하는 함수 
                        mCurrentRow = mCurrentRow + 1;              // 인쇄 라인 증가 
                        vPageRowCount = vPageRowCount + 1;          // 인쇄 라인 카운트 

                        //vWorkCenter_Code = iConv.ISNull(vRow_L["WORKCENTER_CODE"]);  //작업장코드 

                        if (vRow == vTotalRow)
                        {
                            //// 마지막 데이터 이면 처리할 사항 기술     
                            ////합계 인쇄 
                            //mCurrentRow = mCurrentRow + ((m1stLastRow - vPageRowCount) * 2);

                            //mPrinting.XLSetCell(mCurrentRow, 22, string.Format("{0:###,###}", vPO_REQ_AMOUNT));
                            //mPrinting.XLSetCell(mCurrentRow + 30, 22, string.Format("{0:###,###}", vPO_REQ_AMOUNT));  //add 
                        }
                        else
                        {
                            // 새로운 페이지 체크 및 생성.
                            IsNewPage(mPageNumber, vPageRowCount);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;    // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = 4;
                            }
                        }
                    }
                    // Page 표시 //
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

        //Adapter 이용 
        public int WriteMain(InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_HEADER
                            , InfoSummit.Win.ControlAdv.ISDataAdapter pIDA_LINE)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;

            int vRow = 0;                   //인쇄되는 Data Row 위치 
            int vTotalRow = 0;              //총 Row 수
            int vPageRowCount = 0;          //Page당 인쇄되는 Row 수 
            decimal vTotalPageNumber = 0;   //총 인쇄 Page 수(올림 함수 parameter Type)           

            string vPrint_Value = string.Empty;
            string vWorkCenter_Code = string.Empty;     //작업장 코드 

            //인쇄 범위 설정//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 45;
            mCopy_EndRow = 34;
            m1stLastRow = 33;                   //첫장 최종 인쇄 라인.

            mDefaultPageRow = 6;                // 페이지 증가후 PageCount 기본값.
            mCurrentRow = 6;                    // 현재 인쇄되는 row 위치.
            vPageRowCount = mCurrentRow;    // 첫장에 대해서는 시작row부터 체크.

            // 인쇄물 종류 : 헤더 - 라인 관계일 경우 //
            try
            {
                //인쇄물 헤더 : 원본 쉬트에 데이터값을 적용 
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

            // 인쇄 : 라인 //
            try
            {
                //총 row수 
                vTotalRow = pIDA_LINE.CurrentRows.Count;
                vTotalPageNumber = Math.Ceiling((iConv.ISDecimaltoZero(vTotalRow) / iConv.ISDecimaltoZero(m1stLastRow)));

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                //---- Line Write Start ----
                if (vTotalRow > 0)
                {
                    // 원본을 복사해서 타깃쉬트에 복사/붙여 넣기//
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
                            //다른 작업장 : 새로운 페이지 체크 및 생성.
                            mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + 1;
                            vPageRowCount = m1stLastRow;

                            IsNewPage(mPageNumber, vPageRowCount);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (m1stLastRow - vPageRowCount) + mDefaultPageRow - 1;    // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }

                        // 작업장 및 바코드 인쇄 //
                        if (vWorkCenter_Code == string.Empty || vWorkCenter_Code != iConv.ISNull(vRow2["WORKCENTER_CODE"]))
                        {
                            //요청정보 
                            vPrint_Value = string.Format("{0}", vRow2["REQ_ORDER_NO"]);
                            mPrinting.XLSetCell((mCurrentRow - 3), 27, vPrint_Value);

                            //출고처 
                            vPrint_Value = string.Format("{0}", vRow2["WORKCENTER_DESCRIPTION"]);
                            mPrinting.XLSetCell((mCurrentRow - 2), 6, vPrint_Value);

                            //출고처 코드
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

                        mCurrentRow = XLLine(vRow2);        // 실제 인쇄하는 함수 

                        vWorkCenter_Code = iConv.ISNull(vRow2["WORKCENTER_CODE"]);  //작업장코드 

                        if (vRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술                                                       
                        }
                        else
                        {
                            // 새로운 페이지 체크 및 생성.
                            IsNewPage(mPageNumber, vPageRowCount);
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + mDefaultPageRow;    // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }

                    //// Page 표시 //
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
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            object vObject = null;
            string vPrintValue = string.Empty;
            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 만기 일자 기간
                vObject = string.Format("{0} ~ {1}", pDUE_DATE_FR, pDUE_DATE_TO);
                vPrintValue = iConv.ISNull(vObject);
                vXLine = 3;
                vXLColumn = 14;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintValue); 

                //인쇄 일시
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

        //전달받은 변수를 이용한 인쇄 
        private void XLHeader(System.Data.DataRow pRow)
        {// 헤더 인쇄.
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

                //출력일시 
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
            //pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.

            int vXLine = mCurrentRow; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 1;

            // 인쇄 변수 //
            object vObject = null;
            string vPrint_Value = string.Empty;

            try
            { // 타겟쉬트 Active.
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
            //pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.

            int vXLine = mCurrentRow; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 1;

            // 인쇄 변수 //
            object vObject = null;
            string vPrint_Value = string.Empty;

            try
            { // 타겟쉬트 Active.
                mPrinting.XLActiveSheet(mTargetSheet);

                //ITEM CODE 
                vObject = pRow["PRT_ITEM_CODE"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 1;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //자재명                 
                vObject = pRow["PRT_ITEM_DESCRIPTION"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 5;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //자재규격                 
                vObject = pRow["PRT_ITEM_SPECIFICATION"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 12;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //UOM                   
                vObject = pRow["REQ_UOM_CODE"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 21;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);

                //요청 잔량 
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

                //순서 
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

                //박스번호                 
                vObject = pRow["ONHAND_BOX_NO"];
                vPrint_Value = string.Format("{0}", vObject);
                vXLColumn = 27;
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrint_Value);


                //재고량                 
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

                //PICKING 량                 
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
        //{// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
        //    int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // 원본을 복사해서 타겟 에 복사해 넣음.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //합계
        //        vXLColumnIndex = 2;
        //        mPrinting.XLCellMerge(pXLine, vXLColumnIndex, pXLine, 35, true);
        //        vConvertString = string.Format("{0}", "계");
        //        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

        //        //17 - 건수
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

        //        //31 - 공급가액
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
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.

            int vXLRow = 33; //엑셀에 내용이 표시되는 행 번호
            int vXLCol = 46;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
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
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.

            int vXLRow = 36; //엑셀에 내용이 표시되는 행 번호
            int vXLCol = 22;
            
            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
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
            { // pPrintingLine : 현재 출력된 행.                
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

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            // page수 표시.
            mPageNumber = mPageNumber + 1;

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

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
