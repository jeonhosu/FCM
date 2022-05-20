using System;
using ISCommonUtil;

namespace FCMF0575
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SOURCE1";
        private string mSourceSheet2 = "SOURCE2";
        
        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.
        
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 1;
        private int mCopy_StartRow = 1;
        private int mCopy_EndCol = 72;
        private int mCopy_EndRow = 40;

        private int m1stLastRow = 39;       //첫장 최종 인쇄 라인.
        
        private int mPrintingLastRow = 39;  //최종 인쇄 라인 다음 라인.

        private int mDefaultPageRow = 4;
        private int mCurrentRow = 5;       //현재 인쇄되는 row 위치.
        
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

        #region ----- Array Set 1 (총합계)----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("CUSTOMER_COUNT");
            pGDColumn[1] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("GL_AMOUNT_5");
            pGDColumn[3] = pGrid.GetColumnToIndex("GL_AMOUNT_4");
            pGDColumn[4] = pGrid.GetColumnToIndex("GL_AMOUNT_3");
            pGDColumn[5] = pGrid.GetColumnToIndex("GL_AMOUNT_2");
            pGDColumn[6] = pGrid.GetColumnToIndex("GL_AMOUNT_1");
            pGDColumn[7] = pGrid.GetColumnToIndex("VAT_AMOUNT_5");
            pGDColumn[8] = pGrid.GetColumnToIndex("VAT_AMOUNT_4");
            pGDColumn[9] = pGrid.GetColumnToIndex("VAT_AMOUNT_3");
            pGDColumn[10] = pGrid.GetColumnToIndex("VAT_AMOUNT_2");
            pGDColumn[11] = pGrid.GetColumnToIndex("VAT_AMOUNT_1");
                        
            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 19;
            pXLColumn[1] = 24;
            pXLColumn[2] = 28;
            pXLColumn[3] = 31;
            pXLColumn[4] = 34;
            pXLColumn[5] = 37;
            pXLColumn[6] = 40;
            pXLColumn[7] = 42;
            pXLColumn[8] = 45;
            pXLColumn[9] = 48;
            pXLColumn[10] = 51;
            pXLColumn[11] = 54;
        }

        #endregion;

        #region ----- Array Set 2 (명세) -----

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[15];
            pXLColumn = new int[15];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = 0;
            pGDColumn[1] = pGrid.GetColumnToIndex("TAX_REG_NO");
            pGDColumn[2] = pGrid.GetColumnToIndex("CUSTOMER_DESC");
            pGDColumn[3] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[4] = pGrid.GetColumnToIndex("GL_AMOUNT_5");
            pGDColumn[5] = pGrid.GetColumnToIndex("GL_AMOUNT_4");
            pGDColumn[6] = pGrid.GetColumnToIndex("GL_AMOUNT_3");
            pGDColumn[7] = pGrid.GetColumnToIndex("GL_AMOUNT_2");
            pGDColumn[8] = pGrid.GetColumnToIndex("GL_AMOUNT_1");
            pGDColumn[9] = pGrid.GetColumnToIndex("VAT_AMOUNT_5");
            pGDColumn[10] = pGrid.GetColumnToIndex("VAT_AMOUNT_4");
            pGDColumn[11] = pGrid.GetColumnToIndex("VAT_AMOUNT_3");
            pGDColumn[12] = pGrid.GetColumnToIndex("VAT_AMOUNT_2");
            pGDColumn[13] = pGrid.GetColumnToIndex("VAT_AMOUNT_1");
            pGDColumn[14] = 0;


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 3;
            pXLColumn[1] = 6;
            pXLColumn[2] = 12;
            pXLColumn[3] = 22;
            pXLColumn[4] = 25;
            pXLColumn[5] = 28;
            pXLColumn[6] = 31;
            pXLColumn[7] = 34;
            pXLColumn[8] = 37;
            pXLColumn[9] = 39;
            pXLColumn[10] = 42;
            pXLColumn[11] = 45;
            pXLColumn[12] = 48;
            pXLColumn[13] = 51;
            pXLColumn[14] = 53;
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

        #region ----- Header Write Method ----

        public void HeaderWrite(string pSort_Type, string pBalance_Date, InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID, object pPRINT_DATE)
        {// 헤더 인쇄.
            int vXLine = 3;
            int vXLColumn = 0;
             
            try
            {
                if (pSort_Type == "ACCOUNT")
                {
                    mPrinting.XLActiveSheet(mSourceSheet1);
                }
                else
                {
                    mPrinting.XLActiveSheet(mSourceSheet2);
                }

                // 기준일
                vXLine = 3;
                vXLColumn = 67;
                mPrinting.XLSetCell(vXLine, vXLColumn, pBalance_Date);
                  
                //기준일자
                vXLine = 40;
                vXLColumn = 55;
                mPrinting.XLSetCell(vXLine, vXLColumn, string.Format("[{0}]", pPRINT_DATE)); 
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;
        
        #region ----- Excel Line Wirte Methods ----

        public int LineWrite_Account(InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty; 

            int vTotalRow = pGRID.RowCount;
            int vPageRowCount = 4; 

            //인쇄 영역 설정//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 72;
            mCopy_EndRow = 40;

            m1stLastRow = 39;       //첫장 최종 인쇄 라인.            
            mPrintingLastRow = 39;  //최종 인쇄 라인 다음 라인.
                        
            mCurrentRow = 5;       //현재 인쇄되는 row 위치. 

            try
            {
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓. 

                if (vTotalRow > 0)
                {
                    #region ----- Write Page Copy(SourceSheet => TargetSheet) ----
                                        
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                    #endregion;


                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {//Row

                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_Account(vRow, mCurrentRow, pGRID); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow)
                        {
                             
                        }
                        else
                        {
                            IsNewPage(vPageRowCount, mSourceSheet1);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow + mDefaultPageRow);  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
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

        public int LineWrite_Vendor(InfoSummit.Win.ControlAdv.ISGridAdvEx pGRID)
        {// 실제 호출되는 부분.
            string vMessage = string.Empty;

            int vTotalRow = pGRID.RowCount;
            int vPageRowCount = 4;

            //인쇄 영역 설정//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 72;
            mCopy_EndRow = 40;

            m1stLastRow = 39;       //첫장 최종 인쇄 라인.            
            mPrintingLastRow = 39;  //최종 인쇄 라인 다음 라인.

            mCurrentRow = 5;       //현재 인쇄되는 row 위치. 

            try
            {
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓. 

                if (vTotalRow > 0)
                {
                    #region ----- Write Page Copy(SourceSheet => TargetSheet) ----

                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, 1);

                    #endregion;


                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {//Row
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_Vendor(vRow, mCurrentRow, pGRID); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow)
                        {

                        }
                        else
                        {
                            IsNewPage(vPageRowCount, mSourceSheet2);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow + mDefaultPageRow);  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
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

        #region ----- Excel Write [KRW] Method -----
         
        private int LineWrite_Account(int pRow, int pXLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는  
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty; 

            try
            {                
                //예정일
                vXLColumn = 1;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DUE_DATE"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계정코드
                vXLColumn = 5;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("ACCOUNT_CODE"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계정명.
                vXLColumn = 9;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("ACCOUNT_DESC"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //거래처명.
                vXLColumn = 18;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("VENDOR_NAME"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //확정.
                vXLColumn = 29;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("CONFIRM_FLAG")); 
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전표일자.
                vXLColumn = 31;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_DATE"));  
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //통화.
                vXLColumn = 35;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("CURRENCY_CODE"));  
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //원화금액.
                vXLColumn = 37;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_AMOUNT"));  
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외화금액.
                vXLColumn = 43;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("CURR_GL_AMOUNT"));  
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###.####}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전표번호
                vXLColumn = 49;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_NUM"));  
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //적요.
                vXLColumn = 55;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("REMARK"));   
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        private int LineWrite_Vendor(int pRow, int pXLine, InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는  
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
                //예정일
                vXLColumn = 1;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("DUE_DATE"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //거래처명.
                vXLColumn = 5;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("VENDOR_NAME"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계정코드
                vXLColumn = 16;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("ACCOUNT_CODE"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //계정명.
                vXLColumn = 20;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("ACCOUNT_DESC"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);                

                //확정.
                vXLColumn = 29;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("CONFIRM_FLAG"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전표일자.
                vXLColumn = 31;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_DATE"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //통화.
                vXLColumn = 35;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("CURRENCY_CODE"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //원화금액.
                vXLColumn = 37;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_AMOUNT"));
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //외화금액.
                vXLColumn = 43;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("CURR_GL_AMOUNT"));
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###.####}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //전표번호
                vXLColumn = 49;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("GL_NUM"));
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //적요.
                vXLColumn = 55;
                vObject = pGrid.GetCellValue(pRow, pGrid.GetColumnToIndex("REMARK"));
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        private int LineWrite1(System.Data.DataRow pRow, int pXLine, bool pPrint_Flag)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;
            //decimal vConvertDecimal = 0m;

            try
            {
                if (pPrint_Flag == true)
                {
                    //계정명
                    vXLColumn = 1;
                    vObject = pRow["ACCOUNT_DESC"];
                    if (iString.ISNull(vObject) != string.Empty)
                    {
                        vConvertString = string.Format("{0}", vObject);

                    }
                    else
                    {
                        vConvertString = string.Empty;
                    }
                    mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);
                }
                //거래처코드
                vXLColumn = 5;
                vObject = pRow["CUSTOMER_CODE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //거래처 사업자번호.
                vXLColumn = 9;
                vObject = pRow["TAX_REG_NO"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //거래처 명.
                vXLColumn = 14;
                vObject = pRow["CUSTOMER_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //잔액월.
                vXLColumn = 22;
                vObject = pRow["BALANCE_MONTH"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);

                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //금액.
                vXLColumn = 24;
                vObject = pRow["GL_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //현금지급액.
                vXLColumn = 29;
                vObject = pRow["CASH_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //신한금액
                vXLColumn = 34;
                vObject = pRow["BILL_88_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //신한 어음만기일.
                vXLColumn = 39;
                vObject = pRow["DUE_88_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //우리금액
                vXLColumn = 42;
                vObject = pRow["BILL_20_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //우리 어음만기일.
                vXLColumn = 47;
                vObject = pRow["DUE_20_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //국민금액
                vXLColumn = 50;
                vObject = pRow["BILL_04_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //국민 어음만기일.
                vXLColumn = 55;
                vObject = pRow["DUE_04_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기업금액
                vXLColumn = 58;
                vObject = pRow["BILL_03_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //기업 어음만기일.
                vXLColumn = 63;
                vObject = pRow["DUE_03_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //씨티금액
                vXLColumn = 66;
                vObject = pRow["BILL_53_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //씨티 어음만기일.
                vXLColumn = 71;
                vObject = pRow["DUE_53_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vConvertString = string.Format("{0}", vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //-------------------------------------------------------------------
                vXLine++;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        #endregion;

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount, string pActiveSheet)
        {
            int iDefaultEndRow = 1;
            if (mPageNumber == 1)
            {
                if (pPageRowCount == m1stLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - m1stLastRow;
                    mCopyLineSUM = CopyAndPaste(mPrinting, pActiveSheet, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
            else
            {
                if (pPageRowCount == mPrintingLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    iDefaultEndRow = mCopy_EndRow - mPrintingLastRow;
                    mCopyLineSUM = CopyAndPaste(mPrinting, pActiveSheet, mCurrentRow + iDefaultEndRow);
                }
                else
                {
                    mIsNewPage = false;
                }
            }
        }

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNumber(string pActiveSheet, object pPageNumber)
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.
            //int vXLRow = 51; //엑셀에 내용이 표시되는 행 번호
            //int vXLCol = 51;
            //if (iString.ISDecimaltoZero(pPageNumber) == 1)
            //{//첫장 적용
            //    vXLRow = 51; //엑셀에 내용이 표시되는 행 번호
            //    vXLCol = 52;
            //}
            //else
            //{//첫장 외.
            //    vXLRow = 54; //엑셀에 내용이 표시되는 행 번호
            //    vXLCol = 51;
            //}

            //try
            //{ // 원본을 복사해서 타겟 에 복사해 넣음.(
            //    mPrinting.XLActiveSheet(pActiveSheet);
            //    mPrinting.XLSetCell(vXLRow, vXLCol, pPageNumber);
            //}
            //catch (System.Exception ex)
            //{
            //    mMessageError = ex.Message;
            //    mAppInterface.OnAppMessageEvent(mMessageError);
            //    System.Windows.Forms.Application.DoEvents();
            //}
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            // page수 표시.
            mPageNumber = mPageNumber + 1;
            XLPageNumber(pActiveSheet, mPageNumber);

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


            //int vCopySumPrintingLine = pCopySumPrintingLine;

            //int vCopyPrintingRowSTART = vCopySumPrintingLine;
            //vCopySumPrintingLine = vCopySumPrintingLine + mIncrementCopyMAX;
            //int vCopyPrintingRowEnd = vCopySumPrintingLine;

            //pPrinting.XLActiveSheet("SourceTab1");
            //object vRangeSource = pPrinting.XLGetRange(mCopyColumnSTART, 1, mIncrementCopyMAX, mCopyColumnEND); //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLActiveSheet("Destination");
            //object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, 1, vCopyPrintingRowEnd, mCopyColumnEND); //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            //pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            //mPageNumber++; //페이지 번호
            //// 페이지 번호 표시.
            ////string vPageNumberText = string.Format("Page {0}/{1}", mPageNumber, mPageTotalNumber);
            ////int vRowSTART = vCopyPrintingRowEnd - 2;
            ////int vRowEND = vCopyPrintingRowEnd - 2;
            ////int vColumnSTART = 30;
            ////int vColumnEND = 33;
            ////mPrinting.XLCellMerge(vRowSTART, vColumnSTART, vRowEND, vColumnEND, false);
            ////mPrinting.XLSetCell(vRowSTART, vColumnSTART, vPageNumberText); //페이지 번호, XLcell[행, 열]

            //return vCopySumPrintingLine;
        }

        #endregion;

        #region ----- Printing Methods ----

        public void Printing(int pPageSTART, int pPageEND)
        {
            mPrinting.XLPreviewPrinting(pPageSTART, pPageEND, 1);
        }

        #endregion;

        #region ----- Save Methods ----

        public void SAVE(string pSaveFileName)
        {
            if (iString.ISNull(pSaveFileName) == string.Empty)
            {
                return;
            }

            //int vMaxNumber = MaxIncrement(pSavePath.ToString(), pSaveFileName);
            //vMaxNumber = vMaxNumber + 1;
            //string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            //vSaveFileName = string.Format("{0}\\{1}.xls", pSavePath, vSaveFileName);
            //mPrinting.XLSave(vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
