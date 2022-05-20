using System;
using ISCommonUtil;

namespace FCMF0505
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
        private string mTargetSheet = "Destination";
        private string mSourceSheet1 = "SourceTab1";
        //private string mSourceSheet2 = "SOURCE2";
        
        private string mMessageError = string.Empty;
        private string mXLOpenFileName = string.Empty;

        //private int mPageTotalNumber = 0;
        private int mPageNumber = 0;

        private bool mIsNewPage = false;  // 첫 페이지 체크.
        
        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingStartRow = 0;  //최종 인쇄 라인.
        private int mPrintingLastRow = 0;  //최종 인쇄 라인.

        private int mCurrentRow = 0;       //현재 인쇄되는 row 위치.
        private int mDefaultPageRow = 0;    // 페이지 증가후 PageCount 기본값.
        private int mDefaultPageRow2 = 0;    // 페이지 증가후 PageCount 기본값.

        string vO_Account = null;
        string vN_Account = null;
        string vN_ACCOUNT_DESC = null;
        string vACCOUNT = null;
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

        #region ----- Line SLIP Methods ----

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {
            // 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[14];

            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("GL_DATE");
            pGDColumn[1] = pGrid.GetColumnToIndex("SLIP_NUM");
            pGDColumn[2] = pGrid.GetColumnToIndex("ACCOUNT_CODE");
            pGDColumn[3] = pGrid.GetColumnToIndex("ACCOUNT_DESC");
            pGDColumn[4] = pGrid.GetColumnToIndex("REMARK");
            pGDColumn[5] = pGrid.GetColumnToIndex("SUPP_CUST_CODE");
            pGDColumn[6] = pGrid.GetColumnToIndex("SUPP_CUST_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("DR_AMOUNT");
            pGDColumn[8] = pGrid.GetColumnToIndex("CR_AMOUNT");
            pGDColumn[9] = pGrid.GetColumnToIndex("REMAIN_AMOUNT");
            pGDColumn[10] = pGrid.GetColumnToIndex("CURRENCY_CODE");
            pGDColumn[11] = pGrid.GetColumnToIndex("DR_CURRENCY_AMOUNT");
            pGDColumn[12] = pGrid.GetColumnToIndex("CR_CURRENCY_AMOUNT");
            pGDColumn[13] = pGrid.GetColumnToIndex("REMAIN_CURR_AMOUNT");
           

        }

        private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn)
        {
            // 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[14];

            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("GL_DATE");
            pGDColumn[1] = pGrid.GetColumnToIndex("SLIP_NUM");
            pGDColumn[2] = pGrid.GetColumnToIndex("ACCOUNT_CODE");
            pGDColumn[3] = pGrid.GetColumnToIndex("ACCOUNT_DESC");
            pGDColumn[4] = pGrid.GetColumnToIndex("REMARK");
            pGDColumn[5] = pGrid.GetColumnToIndex("SUPP_CUST_CODE");
            pGDColumn[6] = pGrid.GetColumnToIndex("SUPP_CUST_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("DR_AMOUNT");
            pGDColumn[8] = pGrid.GetColumnToIndex("CR_AMOUNT");
            pGDColumn[9] = pGrid.GetColumnToIndex("REMAIN_AMOUNT");
            pGDColumn[10] = pGrid.GetColumnToIndex("CURRENCY_CODE");
            pGDColumn[11] = pGrid.GetColumnToIndex("DR_CURRENCY_AMOUNT");
            pGDColumn[12] = pGrid.GetColumnToIndex("CR_CURRENCY_AMOUNT");
            pGDColumn[13] = pGrid.GetColumnToIndex("REMAIN_CURR_AMOUNT");
            //pGDColumn[14] = pGrid.GetColumnToIndex("ACCOUNT_GROUP_CODE");
            //pGDColumn[15] = pGrid.GetColumnToIndex("ACCOUNT_GROUP_DESC");

        }

        #endregion;

        #region ----- Array Set 2 ----

        //private void SetArray2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        //{// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
        //    pGDColumn = new int[3];
        //    pXLColumn = new int[3];
        //    // 그리드 or 아답터 위치.
        //    pGDColumn[0] = pGrid.GetColumnToIndex("VAT_COUNT");
        //    pGDColumn[1] = pGrid.GetColumnToIndex("GL_AMOUNT");
        //    pGDColumn[2] = pGrid.GetColumnToIndex("VAT_AMOUNT");

        //    // 엑셀에 인쇄해야 할 위치.
        //    pXLColumn[0] = 20;
        //    pXLColumn[1] = 25;
        //    pXLColumn[2] = 30;
        //}

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

        public void HeaderWrite_1(object pPERIOD, object pACCOUNT_DESC)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintDateTime = string.Format("[{0} {1}]", vPrintingDate, vPrintingTime);

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);

                // 기간
                vXLine = 4;
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPERIOD);

                vXLine = 4;
                vXLColumn = 56;
                mPrinting.XLSetCell(vXLine, vXLColumn, pACCOUNT_DESC);

                //vXLine = 4;
                //vXLColumn = 56;
                //mPrinting.XLSetCell(vXLine, vXLColumn, vPrintDateTime);
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite_2(object pPERIOD, object pACCOUNT_DESC)
        {// 헤더 인쇄.
            int vXLine = 0;
            int vXLColumn = 0;

            string vPrintingDate = System.DateTime.Now.ToString("yyyy-MM-dd", null);
            string vPrintingTime = System.DateTime.Now.ToString("HH:mm:ss", null);
            string vPrintDateTime = string.Format("[{0} {1}]", vPrintingDate, vPrintingTime);

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet1);
                // 기간
                vXLine = 4;
                vXLColumn = 24;
                mPrinting.XLSetCell(vXLine, vXLColumn, pPERIOD);

                vXLine = 4;
                vXLColumn = 56;
                mPrinting.XLSetCell(vXLine, vXLColumn, pACCOUNT_DESC);

               
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

    
        #endregion;

        #region ----- Header1 Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// 헤더 인쇄.
            int vXLine = 9; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                for (int i = 0; i <= pGrid.RowCount; i++)
                {
                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[0];
                    vXLColumnIndex = pXLColumn[0];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[1];
                    vXLColumnIndex = pXLColumn[1];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }

                    // 숫자형 예시.
                    vGDColumnIndex = pGDColumn[2];
                    vXLColumnIndex = pXLColumn[2];
                    vObject = pGrid.GetCellValue(i, vGDColumnIndex);
                    IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                    if (IsConvert == true)
                    {
                        vConvertString = string.Format("{0:##,###,###,###,###,###,###,###,###}", vConvertDecimal);
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    else
                    {
                        vConvertString = string.Empty;
                        mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);
                    }
                    vXLine++;
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

        #region ----- Line Write Method -----

        private int XLLine_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn)
        {
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;
                        
            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vCONVERT_DATE = new DateTime();
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);
                
                //0-회계일자
                vGDColumnIndex = pGDColumn[0];                
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertDate(vObject, out vCONVERT_DATE);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vCONVERT_DATE.ToShortDateString());
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 1;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //1-회계번호
                vGDColumnIndex = pGDColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 5;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);


                //2-계정코드
                vGDColumnIndex = pGDColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 8;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //3-계정명
                vGDColumnIndex = pGDColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 12;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);


               
                //5-적요
                vGDColumnIndex = pGDColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 18;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //6 - 거래처코드
                vGDColumnIndex = pGDColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 29;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //7 - 거래처명
                vGDColumnIndex = pGDColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 32;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //8 - 차변
                vGDColumnIndex = pGDColumn[7];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 42;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //9- 대변금액
                vGDColumnIndex = pGDColumn[8];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 47;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //10- 잔액
                vGDColumnIndex = pGDColumn[9];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 52;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);


                //11- 통화
                vGDColumnIndex = pGDColumn[10];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 57;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //12 - 외화차변
                vGDColumnIndex = pGDColumn[11];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##.###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 59;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //13- 외화대변금액
                vGDColumnIndex = pGDColumn[12];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##.###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 63;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //14- 외화잔액
                vGDColumnIndex = pGDColumn[13];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##.###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 67;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);



                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;

            return pXLine;
        }
        
        private int XLLine_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int pGridRow, int pXLine, int[] pGDColumn)
        {
            // pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

            int vGDColumnIndex = 0;
            int vXLColumnIndex = 0;

            // 사용되는 형식 지정.
            object vObject = null;
            string vConvertString = string.Empty;
            decimal vConvertDecimal = 0m;
            DateTime vCONVERT_DATE = new DateTime();
            bool IsConvert = false;

            try
            { // 원본을 복사해서 타겟 에 복사해 넣음.(
                mPrinting.XLActiveSheet(mTargetSheet);

                //0-회계일자
                vGDColumnIndex = pGDColumn[0];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertDate(vObject, out vCONVERT_DATE);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vCONVERT_DATE.ToShortDateString());
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 1;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //1-회계번호
                vGDColumnIndex = pGDColumn[1];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 5;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);


                //2-계정코드
                vGDColumnIndex = pGDColumn[2];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 8;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //3-계정명
                vGDColumnIndex = pGDColumn[3];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 12;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);



                //5-적요
                vGDColumnIndex = pGDColumn[4];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 18;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //6 - 거래처코드
                vGDColumnIndex = pGDColumn[5];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 29;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //7 - 거래처명
                vGDColumnIndex = pGDColumn[6];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 32;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //8 - 차변
                vGDColumnIndex = pGDColumn[7];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 42;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //9- 대변금액
                vGDColumnIndex = pGDColumn[8];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 47;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //10- 잔액
                vGDColumnIndex = pGDColumn[9];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 52;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);


                //11- 통화
                vGDColumnIndex = pGDColumn[10];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertString(vObject, out vConvertString);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0}", vConvertString);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 57;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //12 - 외화차변
                vGDColumnIndex = pGDColumn[11];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##.###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 59;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //13- 외화대변금액
                vGDColumnIndex = pGDColumn[12];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##.###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 63;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);

                //14- 외화잔액
                vGDColumnIndex = pGDColumn[13];
                vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                IsConvert = IsConvertNumber(vObject, out vConvertDecimal);
                if (IsConvert == true)
                {
                    vConvertString = string.Format("{0:##.###,###,###,###,###,###,###,###}", vConvertDecimal);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                vXLColumnIndex = 67;
                mPrinting.XLSetCell(vXLine, vXLColumnIndex, vConvertString);


                ////2-계정코드
                //vGDColumnIndex = pGDColumn[14];
                //vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                //IsConvert = IsConvertString(vObject, out vConvertString);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumnIndex = 54;
                //mPrinting.XLSetCell(mPrintingStartRow +3 , vXLColumnIndex, vConvertString);

                ////3-계정명
                //vGDColumnIndex = pGDColumn[15];
                //vObject = pGrid.GetCellValue(pGridRow, vGDColumnIndex);
                //IsConvert = IsConvertString(vObject, out vConvertString);
                //if (IsConvert == true)
                //{
                //    vConvertString = string.Format("{0}", vConvertString);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //vXLColumnIndex = 54;
                //mPrinting.XLSetCell(mPrintingStartRow +3, vXLColumnIndex, vConvertString);



                //-------------------------------------------------------------------
                vXLine = vXLine + 1;
                //-------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }

            pXLine = vXLine;

            return pXLine;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        //private int XL_TOTAL_Line(int pXLine, int[] pXLColumn)
        //{// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행. pGDColumn : 그리드 위치, pXLColumn : 엑셀 위치.
        //    int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호

        //    int vXLColumnIndex = 0;

        //    string vConvertString = string.Empty;
        //    decimal vConvertDecimal = 0m;
        //    bool IsConvert = false;

        //    try
        //    { // 원본을 복사해서 타겟 에 복사해 넣음.(
        //        mPrinting.XLActiveSheet(mTargetSheet);

        //        //10 - 보증금
        //        vXLColumnIndex = pXLColumn[10];
        //        IsConvert = IsConvertNumber(mTOT_DEPOSIT_AMOUNT, out vConvertDecimal);
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
        //        //11 - 월임대료
        //        vXLColumnIndex = pXLColumn[11];
        //        IsConvert = IsConvertNumber(mTOT_MONTHLY_RENT_AMOUNT, out vConvertDecimal);
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
        //        //12 - 합계
        //        vXLColumnIndex = pXLColumn[12];
        //        IsConvert = IsConvertNumber(mTOT_LEASE_SUM_AMOUNT, out vConvertDecimal);
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
        //        //13 - 보증금이자
        //        vXLColumnIndex = pXLColumn[13];
        //        IsConvert = IsConvertNumber(mTOT_DEPOSIT_INTEREST_AMT, out vConvertDecimal);
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
        //        //14 - 월임대료(계)
        //        vXLColumnIndex = pXLColumn[14];
        //        IsConvert = IsConvertNumber(mTOT_MONTHLY_RENT_SUM_AMT, out vConvertDecimal);
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

        #endregion;

        #region ----- Excel Wirte [Line] Methods ----

        public int LineWrite_1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid,object pPeriod)
        {// 실제 호출되는 부분.
            mPageNumber = 0;
            string vMessage = string.Empty;
            
            int[] vGDColumn;
            int vTotalRow = 0;
            int vPageRowCount = 0;

            // 인쇄 1장의 최대 인쇄정보.
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 70;
            mCopy_EndRow = 38;
            mPrintingLastRow = 38;  //최종 인쇄 라인.

            mCurrentRow = 8;       //현재 인쇄되는 row 위치.
            mDefaultPageRow = 7;    // 페이지 증가후 PageCount 기본값.
            
            try
            {
                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);
                                
                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    SetArray1(pGrid, out vGDColumn);
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XLLine_1(pGrid, vRow, mCurrentRow, vGDColumn); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vRow == vTotalRow - 1)
                        {
                            //XL_TOTAL_Line(12, vXLColumn);
                        }
                        else
                        {
                            IsNewPage(vPageRowCount, pPeriod, vACCOUNT);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow;
                            }
                        }
                    }
                   // mPrinting.XL_LineClearInSide(mCurrentRow, 1, mCurrentRow+30, 71);

                    mPrinting.XL_LineClearALL(mCurrentRow+1, 1, mCurrentRow + 30, 71);
                    mPrinting.XL_LineDraw_Bottom(mCurrentRow, 1,70,2);
                    

                }
                #endregion;
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mPrinting.XLOpenFileClose();
                mPrinting.XLClose();
            }

            return mPageNumber;
        }
        
        public int LineWrite_2(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid,object pPeriod, object pACCOUNT_DESC)
        {// 실제 호출되는 부분.
            mPageNumber = 0;
            string vMessage = string.Empty;

            int[] vGDColumn;
            int vTotalRow = 0;
            int vPageRowCount = 0;

            // 인쇄 1장의 최대 인쇄정보.
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 70;
            mCopy_EndRow = 38;
            mPrintingStartRow = 1;
            mPrintingLastRow = 38;  //최종 인쇄 라인.

            mCurrentRow = 8;       //현재 인쇄되는 row 위치.
            mDefaultPageRow = 7;    // 페이지 증가후 PageCount 기본값.
            mDefaultPageRow2 = mCopy_EndRow - mCurrentRow +7 ;

            try
            {
                // 실제인쇄되는 행수.
                //int vBy = 35;         
                vTotalRow = pGrid.RowCount;
                vPageRowCount = mCurrentRow - 1;    //첫장에 대해서는 시작row부터 체크.

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.

                // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                //vO_Account = pGrid.GetColumnToIndex("ACCOUNT_CODE").ToString();
                        
                        //pGrid.GetColumnToIndex("ACCOUNT_CODE")


                #region ----- Header Write ----

                //SetArray1(pGrid, out vGDColumn, out vXLColumn);
               

                #endregion;

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    SetArray2(pGrid, out vGDColumn);
                    for (int vRow = 0; vRow < vTotalRow; vRow++)
                    {
                       
                      

                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = XLLine_2(pGrid, vRow, mCurrentRow, vGDColumn); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        int vAccountNum = pGrid.GetColumnToIndex("ACCOUNT_GROUP_CODE");
                         int vAccountDesc = pGrid.GetColumnToIndex("ACCOUNT_GROUP_DESC");
                        vN_Account = pGrid.GetCellValue(vRow+1, vAccountNum).ToString();//pGrid.GetColumnToIndex("ACCOUNT_CODE").ToString();
                        vN_ACCOUNT_DESC = pGrid.GetCellValue(vRow + 1, vAccountDesc).ToString();
                         vACCOUNT = string.Format("({0}){1}", vN_Account, vN_ACCOUNT_DESC);

                        if (vRow <= 1)
                        {
                            vO_Account = vN_Account;

                        }

                        if (vRow == vTotalRow - 1)
                        {

                        }
                        else 
                        {
                            IsNewPage(vPageRowCount, pPeriod, vACCOUNT);   // 새로운 페이지 체크 및 생성.

                         
                            if (mIsNewPage == true)
                            {
                               if(vN_Account == vO_Account)
                               {
                                   mCurrentRow = mCurrentRow + (mCopy_EndRow - mPrintingLastRow) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                   vPageRowCount = mDefaultPageRow;
                               
                               }
                                else
                               {
                                   //mPrinting.XL_LineClearALL(mCurrentRow + 1, 1, mPrintingStartRow - 1, 71);

                                   mCurrentRow = mPrintingStartRow - mCopy_EndRow +mDefaultPageRow;// mCurrentRow + (mPrintingLastRow - mCurrentRow) + mDefaultPageRow + 1;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                //mCurrentRow = mPrintingLastRow- mCurrentRow + mDefaultPageRow;
                                vPageRowCount = mDefaultPageRow;

                               vO_Account = vN_Account;


                               
                             //  mPrinting.XL_LineDraw_Bottom(mCurrentRow, 1, 70, 2);

                                
                               }
                            }
                           
                        }

                        // }
                    }
                    mPrinting.XL_LineClearALL(mCurrentRow + 1, 1, mPrintingStartRow , 71);
                    mPrinting.XL_LineDraw_Bottom(mCurrentRow, 1, 70, 2);
                }
                #endregion;
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

        #region ----- New Page iF Methods ----

        private void IsNewPage(int pPageRowCount, object pPeriod, string vAccount)
        {
            if (pPageRowCount == mPrintingLastRow)
            { // pPrintingLine : 현재 출력된 행.
                mIsNewPage = true;
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCurrentRow );
            }
            else if (vO_Account != vN_Account && vN_Account != "" )
            {
                mIsNewPage = true;

                HeaderWrite_2(pPeriod, vAccount);
                mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mPrintingStartRow);

                mPrinting.XL_LineClearALL(mCurrentRow + 1, 1, mPrintingStartRow-38, 71);
                mPrinting.XL_LineDraw_Bottom(mCurrentRow, 1, 70, 2);
            }
            else
            {
                mIsNewPage = false;
            }
        }

        #endregion;

        //#region ----- New PageSkip iF Methods ----

        //private void IsNewPageSkip(int pPageRowCount)
        //{
        //    if (pPageRowCount == mPrintingLastRow)
        //    { // pPrintingLine : 현재 출력된 행.
        //        mIsNewPage = true;
        //        mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, mCopy_EndRow + mCurrentRow);
        //    }
        //    else
        //    {
        //        mIsNewPage = false;
        //    }
        //}

        //#endregion;


        #region ----- Copy&Paste Sheet Method ----

        //지정한 ActiveSheet의 범위에 대해  페이지 복사
        private int CopyAndPaste(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {

           

            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            mPrintingStartRow = vPasteEndRow;

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);            
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol); 
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.

           // mPrinting.XL_LineClearALL(mCurrentRow + 1, 1, mPrintingStartRow -1, 71);

            mPageNumber++; //페이지 번호
            
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

        private int CopyAndPaste2(XL.XLPrint pPrinting, string pActiveSheet, int pPasteStartRow)
        {
                   

            int vPasteEndRow = pPasteStartRow + mCopy_EndRow;

            mPrintingStartRow = vPasteEndRow;

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pActiveSheet);            
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 
            //엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(pPasteStartRow, mCopy_StartCol, vPasteEndRow, mCopy_EndCol); 
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);  // 복사.


            mPageNumber++; //페이지 번호
            
            return vPasteEndRow;

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
            if (pSaveFileName == string.Empty)
            {
                return;
            }

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
