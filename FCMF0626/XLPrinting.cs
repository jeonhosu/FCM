using System;
using System.Collections.Generic;
using System.Text;
using ISCommonUtil;

namespace FCMF0626
{
    public class XLPrinting
    {
        #region ----- Variables -----

        ISFunction.ISConvert iString = new ISFunction.ISConvert();
        ISFunction.ISDateTime iDate = new ISFunction.ISDateTime();

        private InfoSummit.Win.ControlAdv.ISAppInterface mAppInterface = null;
        private InfoSummit.Win.ControlAdv.ISMessageAdapter mMessageAdapter = null;

        private XL.XLPrint mPrinting = null;

        private string mMessageError = string.Empty;

        // 쉬트명 정의.
        private string mTargetSheet = "Sheet1";
        private string mSourceSheet1 = "SourceTab1";
        private string mSourceSheet2 = "SourceTab2";

        private int mPageNumber = 0;

        private bool mIsNewPage = false;

        private string mXLOpenFileName = string.Empty;

        // 인쇄된 라인에 합계.
        private int mCopyLineSUM = 0;

        // 인쇄 1장의 최대 인쇄정보.
        private int mCopy_StartCol = 0;
        private int mCopy_StartRow = 0;
        private int mCopy_EndCol = 0;
        private int mCopy_EndRow = 0;
        private int mPrintingLastRow = 0;   //최종 인쇄 라인.
        private int m1stPrintingLastRow = 0; 
        private int mCurrentRow = 0;        //현재 인쇄되는 row 위치.
        private int mDefaultEndPageRow = 1; // 페이지 증가후 PageCount 기본값.
        private int mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.

        private int mCountLinePrinting = 0; //엑셀 라인 Seq

        private decimal mSUM_PL_AMOUNT = 0;     //계획 합계
        private decimal mSUM_AMOUNT = 0;        //예산 합계
        private decimal mSUM_GAP_AMOUNT = 0;    //차액 합계  

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

        #region ----- Array Set 0 ----

        private void SetArray0(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[3];
            pXLColumn = new int[3];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("VAT_COUNT");
            pGDColumn[1] = pGrid.GetColumnToIndex("GL_AMOUNT");
            pGDColumn[2] = pGrid.GetColumnToIndex("VAT_AMOUNT");

            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 12;
            pXLColumn[1] = 22;
            pXLColumn[2] = 34;
        }

        #endregion;

        #region ----- Array Set 1 ----

        private void SetArray1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, out int[] pGDColumn, out int[] pXLColumn)
        {// 그리드의 컬럼에 대한 컬럼인덱스 값 산출
            pGDColumn = new int[12];
            pXLColumn = new int[12];
            // 그리드 or 아답터 위치.
            pGDColumn[0] = pGrid.GetColumnToIndex("PERSON_NUM");
            pGDColumn[1] = pGrid.GetColumnToIndex("NAME");
            pGDColumn[2] = pGrid.GetColumnToIndex("REPRE_NUM");
            pGDColumn[3] = pGrid.GetColumnToIndex("DEPT_NAME");
            pGDColumn[4] = pGrid.GetColumnToIndex("FLOOR_NAME");
            pGDColumn[5] = pGrid.GetColumnToIndex("ABIL_NAME");
            pGDColumn[6] = pGrid.GetColumnToIndex("POST_NAME");
            pGDColumn[7] = pGrid.GetColumnToIndex("ORI_JOIN_DATE");
            pGDColumn[8] = pGrid.GetColumnToIndex("JOIN_DATE");
            pGDColumn[9] = pGrid.GetColumnToIndex("RETIRE_DATE");
            pGDColumn[10] = pGrid.GetColumnToIndex("CONTINUE_YEAR");
            pGDColumn[11] = pGrid.GetColumnToIndex("END_SCH_NAME");


            // 엑셀에 인쇄해야 할 위치.
            pXLColumn[0] = 1;
            pXLColumn[1] = 6;
            pXLColumn[2] = 11;
            pXLColumn[3] = 17;
            pXLColumn[4] = 24;
            pXLColumn[5] = 31;
            pXLColumn[6] = 36;
            pXLColumn[7] = 42;
            pXLColumn[8] = 46;
            pXLColumn[9] = 50;
            pXLColumn[10] = 54;
            pXLColumn[11] = 59;
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

        #region ----- Excel Write -----

        #region ----- Header Write Method ----

        public void HeaderWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pHeader,
                                InfoSummit.Win.ControlAdv.ISDataAdapter pApproval_Step, 
                                object pSOB_DESC, object pLOCAL_DATE)
        {// 헤더 인쇄.
            object vObject;
            string vPrintString;

            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2);
                //업체명 
                vXLine = 47;
                vXLColumn = 1;
                vObject = pSOB_DESC;
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //인쇄일시[PRINTED DATE]
                vXLine = 47;
                vXLColumn = 13;
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vObject = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vObject = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject); 


                mPrinting.XLActiveSheet(mSourceSheet1);
                
                //title
                vXLine = 2;
                vXLColumn = 6;
                vObject = pHeader.CurrentRow["TITLE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);
                
                //신철일자
                vXLine = 7;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["REQ_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }                
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    if (iDate.ISDate(vObject) == true)
                //    {
                //        if (iDate.ISGetDate(vObject).ToShortDateString() == "0001-01-01")
                //        {
                //            vPrintString = iString.ISNull(vObject);
                //        }
                //        else
                //        {
                //            vPrintString = iDate.ISGetDate(vObject).ToShortDateString();
                //        }
                //    }
                //    else
                //    {
                //        vPrintString = iString.ISNull(vObject);
                //    }
                //}
                //else
                //{
                //    vPrintString = string.Empty;
                //}
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //신청부서  
                vXLine = 8;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["BUDGET_DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //기안자[PERSON_NAME]
                vXLine = 9;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["REQ_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //신청번호[BUDGET_ADD_NUM]
                vXLine = 10;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["BUDGET_ADD_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //결재라인.
                //기안자 TITLE.
                vXLine = 5;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["TITLE_10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //기안자.
                vXLine = 6;
                vXLColumn = 14; 
                vObject = pApproval_Step.CurrentRow["NAME_10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //팀장.
                vXLine = 5;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["TITLE_20"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 6;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["NAME_20"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //본부장.
                vXLine = 5;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["TITLE_30"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 6;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["NAME_30"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //기획팀.
                vXLine = 8;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["TITLE_40"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 9;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["NAME_40"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //재경팀
                vXLine = 8;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["TITLE_50"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 9;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["NAME_50"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //대표이사.
                vXLine = 8;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["TITLE_60"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 9;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["NAME_60"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //제목[REMARK]
                vXLine = 13;
                vXLColumn = 1; 
                string vContent = string.Empty;
                vObject = pHeader.CurrentRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    bool isConvert = vObject is string;
                    if (isConvert == true)
                    {
                        vContent = vObject as string;
                        bool isNull = string.IsNullOrEmpty(vContent.Trim());
                        if (isNull != true)
                        {
                            //byte b_CR_Character = 0x0d; //CR
                            //byte b_SP_Character = 0x20; //Space
                            //char vCharOld = (char)b_CR_Character;
                            //char vCharNew = (char)b_SP_Character;
                            //vContent = vContent.Replace(vCharOld, vCharNew);
                            vContent = vContent.Replace("\r", "");
                        }
                    }
                    vPrintString = vContent.ToString(); ;
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //업체명 
                vXLine = 47;
                vXLColumn = 1;
                vObject = pSOB_DESC;
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString); 

                //인쇄일시[PRINTED DATE]
                vXLine = 47;
                vXLColumn = 13;
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vObject = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vObject = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public void HeaderWrite_Etc(InfoSummit.Win.ControlAdv.ISDataAdapter pHeader,
                                    InfoSummit.Win.ControlAdv.ISDataAdapter pApproval_Step,
                                    object pSOB_DESC, object pLOCAL_DATE)
        {// 헤더 인쇄.
            object vObject;
            string vPrintString;

            int vXLine = 0;
            int vXLColumn = 0;

            try
            {
                mPrinting.XLActiveSheet(mSourceSheet2);
                //업체명 
                vXLine = 47;
                vXLColumn = 1;
                vObject = pSOB_DESC;
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //인쇄일시[PRINTED DATE]
                vXLine = 47;
                vXLColumn = 13;
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vObject = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vObject = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);


                mPrinting.XLActiveSheet(mSourceSheet1);

                //title
                vXLine = 2;
                vXLColumn = 6;
                vObject = pHeader.CurrentRow["TITLE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //신철일자
                vXLine = 7;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["REQ_DATE"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    if (iDate.ISDate(vObject) == true)
                //    {
                //        if (iDate.ISGetDate(vObject).ToShortDateString() == "0001-01-01")
                //        {
                //            vPrintString = iString.ISNull(vObject);
                //        }
                //        else
                //        {
                //            vPrintString = iDate.ISGetDate(vObject).ToShortDateString();
                //        }
                //    }
                //    else
                //    {
                //        vPrintString = iString.ISNull(vObject);
                //    }
                //}
                //else
                //{
                //    vPrintString = string.Empty;
                //}
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

                //신청부서  
                vXLine = 8;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["BUDGET_DEPT_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //기안자[PERSON_NAME]
                vXLine = 9;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["REQ_NAME"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //신청번호[BUDGET_ADD_NUM]
                vXLine = 10;
                vXLColumn = 4;
                vObject = pHeader.CurrentRow["BUDGET_ADD_NUM"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //결재라인.
                //기안자 TITLE.
                vXLine = 5;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["TITLE_10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //기안자.
                vXLine = 6;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["NAME_10"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //팀장.
                vXLine = 5;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["TITLE_20"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 6;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["NAME_20"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //본부장.
                vXLine = 5;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["TITLE_30"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 6;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["NAME_30"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //기획팀.
                vXLine = 8;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["TITLE_40"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 9;
                vXLColumn = 14;
                vObject = pApproval_Step.CurrentRow["NAME_40"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //재경팀
                vXLine = 8;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["TITLE_50"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 9;
                vXLColumn = 16;
                vObject = pApproval_Step.CurrentRow["NAME_50"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //대표이사.
                vXLine = 8;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["TITLE_60"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                vXLine = 9;
                vXLColumn = 18;
                vObject = pApproval_Step.CurrentRow["NAME_60"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //제목[REMARK]
                vXLine = 13;
                vXLColumn = 1;
                string vContent = string.Empty;
                vObject = pHeader.CurrentRow["REMARK"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    bool isConvert = vObject is string;
                    if (isConvert == true)
                    {
                        vContent = vObject as string;
                        bool isNull = string.IsNullOrEmpty(vContent.Trim());
                        if (isNull != true)
                        {
                            //byte b_CR_Character = 0x0d; //CR
                            //byte b_SP_Character = 0x20; //Space
                            //char vCharOld = (char)b_CR_Character;
                            //char vCharNew = (char)b_SP_Character;
                            //vContent = vContent.Replace(vCharOld, vCharNew);
                            vContent = vContent.Replace("\r", "");
                        }
                    }
                    vPrintString = vContent.ToString(); ;
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //경영기획팀[REMARK]
                vXLine = 23;
                vXLColumn = 1;
                vContent = string.Empty;
                vObject = pHeader.CurrentRow["REMARK_1"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    bool isConvert = vObject is string;
                    if (isConvert == true)
                    {
                        vContent = vObject as string;
                        bool isNull = string.IsNullOrEmpty(vContent.Trim());
                        if (isNull != true)
                        {
                            //byte b_CR_Character = 0x0d; //CR
                            //byte b_SP_Character = 0x20; //Space
                            //char vCharOld = (char)b_CR_Character;
                            //char vCharNew = (char)b_SP_Character;
                            //vContent = vContent.Replace(vCharOld, vCharNew);
                            vContent = vContent.Replace("\r", "");
                        }
                    }
                    vPrintString = vContent.ToString(); ;
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //재경팀[REMARK_2]
                vXLine = 29;
                vXLColumn = 1;
                vContent = string.Empty;
                vObject = pHeader.CurrentRow["REMARK_2"];
                if (iString.ISNull(vObject) != string.Empty)
                {
                    bool isConvert = vObject is string;
                    if (isConvert == true)
                    {
                        vContent = vObject as string;
                        bool isNull = string.IsNullOrEmpty(vContent.Trim());
                        if (isNull != true)
                        {
                            //byte b_CR_Character = 0x0d; //CR
                            //byte b_SP_Character = 0x20; //Space
                            //char vCharOld = (char)b_CR_Character;
                            //char vCharNew = (char)b_SP_Character;
                            //vContent = vContent.Replace(vCharOld, vCharNew);
                            vContent = vContent.Replace("\r", "");
                        }
                    }
                    vPrintString = vContent.ToString(); ;
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //업체명 
                vXLine = 47;
                vXLColumn = 1;
                vObject = pSOB_DESC;
                if (iString.ISNull(vObject) != string.Empty)
                {
                    vPrintString = string.Format("{0}", vObject);
                }
                else
                {
                    vPrintString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vPrintString);

                //인쇄일시[PRINTED DATE]
                vXLine = 47;
                vXLColumn = 13;
                if (iDate.ISDate(pLOCAL_DATE) == true)
                {
                    vObject = string.Format("[{0:yyyy-MM-dd hh:mm:dd}]", pLOCAL_DATE);
                }
                else
                {
                    vObject = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vObject);

            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
                System.Windows.Forms.Application.DoEvents();
            }
        }

        #endregion;

        #region ----- Header1 (합계) Write Method ----

        private void XLHeader1(InfoSummit.Win.ControlAdv.ISGridAdvEx pGrid, int[] pGDColumn, int[] pXLColumn)
        {// 헤더 인쇄.
            int vXLine = 0; //엑셀에 내용이 표시되는 행 번호

            int vIDX_VAT_TYPE = pGrid.GetColumnToIndex("VAT_TYPE");
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

                for (int i = 0; i < pGrid.RowCount; i++)
                {
                    // 총합계 구분에 따라 인쇄 ROW 지정.
                    if ("T" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//총합계
                        vXLine = 9;
                    }
                    else if ("3" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//신용카드.
                        vXLine = 13;
                    }
                    else if ("11" == iString.ISNull(pGrid.GetCellValue(i, vIDX_VAT_TYPE)))
                    {//현금영수증.
                        vXLine = 10;
                    }

                    //0 - 거래건수.
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
                    //1 - 공급가액
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
                    //2 - 세액
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

        #region ----- Excel Write [Line] Method -----

        private int LineWrite(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty; 

            try
            {
                //[ACCOUNT_CODE]
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

                ////[EXPENDITURE_DATE]
                //vXLColumn = 6;
                //vObject = pRow["EXPENDITURE_DATE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vConvertString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[PL_AMOUNT]
                vXLColumn = 7;
                vObject = pRow["PL_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mSUM_PL_AMOUNT = mSUM_PL_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[AMOUNT]
                vXLColumn = 10;
                vObject = pRow["AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mSUM_AMOUNT = mSUM_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[GAP_AMOUNT]
                vXLColumn = 13;
                vObject = pRow["GAP_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mSUM_GAP_AMOUNT = mSUM_GAP_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[DESCRIPTION]
                vXLColumn = 16;
                vObject = pRow["DESCRIPTION"];
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
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }


        private int LineWrite_Etc(System.Data.DataRow pRow, int pXLine)
        {// pGridRow : 그리드의 현재 읽는 행, pXLine : 엑셀의 인쇄해야 하는 행
            int vXLine = pXLine; //엑셀에 내용이 표시되는 행 번호
            int vXLColumn = 0;

            object vObject = null;
            string vConvertString = string.Empty;

            try
            {
                //[ACCOUNT_CODE]
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

                ////[EXPENDITURE_DATE]
                //vXLColumn = 6;
                //vObject = pRow["EXPENDITURE_DATE"];
                //if (iString.ISNull(vObject) != string.Empty)
                //{
                //    vConvertString = string.Format("{0}", vObject);
                //}
                //else
                //{
                //    vConvertString = string.Empty;
                //}
                //mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[PL_AMOUNT]
                vXLColumn = 7;
                vObject = pRow["BUDGET_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mSUM_PL_AMOUNT = mSUM_PL_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[AMOUNT]
                vXLColumn = 10;
                vObject = pRow["AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mSUM_AMOUNT = mSUM_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[GAP_AMOUNT]
                vXLColumn = 13;
                vObject = pRow["TOTAL_AMOUNT"];
                if (iString.ISDecimal(vObject) == true)
                {
                    vConvertString = string.Format("{0:###,###,###,###,###,###,###}", vObject);
                    mSUM_GAP_AMOUNT = mSUM_GAP_AMOUNT + iString.ISDecimaltoZero(vObject);
                }
                else
                {
                    vConvertString = string.Empty;
                }
                mPrinting.XLSetCell(vXLine, vXLColumn, vConvertString);

                //[DESCRIPTION]
                vXLColumn = 16;
                vObject = pRow["DESCRIPTION"];
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
                //--------------------------------------------------------------------------------------------------
            }
            catch (System.Exception ex)
            {
                mMessageError = ex.Message;
                mAppInterface.OnAppMessageEvent(mMessageError);
            }
            return vXLine;
        }

        #endregion;

        #region ----- TOTAL AMOUNT Write Method -----

        private void SumWrite(int pPrintingLine)
        {
            mPrinting.XLActiveSheet(mTargetSheet);

            //PageNumber 인쇄//
            int vPageCount = 47;
            int vLINE = mPageNumber * mCopy_EndRow;
            for (int r = 1; r <= mPageNumber; r++)
            {
                vLINE = vPageCount * r;
                mPrinting.XLSetCell(vLINE, 7, string.Format("Page {0} of {1}", r, mPageNumber));

                if (r == mPageNumber)
                {
                    //
                }
                else
                {
                    vLINE = vLINE - 1;
                    mPrinting.XLSetCell(vLINE, 1, "");
                }
            }

            //합계 인쇄//
            vLINE = mPageNumber * mCopy_EndRow;
            vLINE = vLINE - 1;
            //mPrinting.XLSetCell(vLINE, 1, "SUM");
            string vAmount = string.Empty;

            //[합계]
            if (mPageNumber == 1)
            {
                vLINE = 46;
                mPrinting.XLSetCell(vLINE, 1, "[총    계]");

                //BACK COLOR.
                mPrinting.XLCellColorBrush(vLINE, 7, vLINE, 13, System.Drawing.Color.Silver);

                //계획합계
                vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_PL_AMOUNT);
                mPrinting.XLSetCell(vLINE, 7, vAmount);

                //예산합계
                vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mSUM_AMOUNT);
                mPrinting.XLSetCell(vLINE, 10, vAmount);

                //차액합계
                vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_GAP_AMOUNT);
                mPrinting.XLSetCell(vLINE, 13, vAmount);

                //XlLineClear(pPrintingLine);

            }
            else
            {
                mPrinting.XLSetCell(vLINE, 1, "[총    계]");

                //BACK COLOR.
                mPrinting.XLCellColorBrush(vLINE, 7, vLINE, 13, System.Drawing.Color.Silver);

                //계획합계
                vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_PL_AMOUNT);
                mPrinting.XLSetCell(vLINE, 7, vAmount);

                //예산합계
                vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###.####}", mSUM_AMOUNT);
                mPrinting.XLSetCell(vLINE, 10, vAmount);

                //차액합계
                vAmount = string.Format("{0:#,###,###,###,###,###,###,###,###,##0}", mSUM_GAP_AMOUNT);
                mPrinting.XLSetCell(vLINE, 13, vAmount);

                //XlLineClear(pPrintingLine);
            } 
        }

        #endregion;

        #region ----- PageNumber Write Method -----

        private void XLPageNumber(string pActiveSheet, object pPageNumber)
        {// 페이지수를 원본쉬트 복사하기 전에 원본쉬트에 기록하고 쉬트를 복사한다.

            int vXLRow = 31; //엑셀에 내용이 표시되는 행 번호
            int vXLCol = 40;

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

        #endregion;

        #endregion;

        #region ----- Excel Wirte MAIN Methods ----

        public int ExcelWrite(InfoSummit.Win.ControlAdv.ISDataAdapter pHeader, InfoSummit.Win.ControlAdv.ISDataAdapter pLine, 
                                InfoSummit.Win.ControlAdv.ISDataAdapter pApproval_Step, 
                                object pSOB_DESC, object pLOCAL_DATE)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            //초기화//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 19;
            mCopy_EndRow = 47;

            mDefaultEndPageRow = 1;
            mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 45;  //최종 인쇄 라인.
            m1stPrintingLastRow = 45;

            mCurrentRow = 31;
            
            // 합계.
            mSUM_PL_AMOUNT = 0;
            mSUM_AMOUNT = 0;
            mSUM_GAP_AMOUNT = 0;

            int vTotalRow = 0;
            int vPageRowCount = 0;  //인쇄후 해당 라인 증가 위해.
            int vLIneRow = 0;
            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pHeader.CurrentRows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    //실제 라인 row수.
                    vTotalRow = pLine.CurrentRows.Count;

                    HeaderWrite(pHeader, pApproval_Step, pSOB_DESC, pLOCAL_DATE);

                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                    //첫장에 대해서는 시작row부터 체크.
                    vPageRowCount = mCurrentRow - 1;   
                     
                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pLine.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            SumWrite(mCurrentRow);      //합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow - 1;
                            }
                        }
                    }
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


        public int ExcelWrite_Etc(InfoSummit.Win.ControlAdv.ISDataAdapter pHeader, InfoSummit.Win.ControlAdv.ISDataAdapter pLine,
                                InfoSummit.Win.ControlAdv.ISDataAdapter pApproval_Step,
                                object pSOB_DESC, object pLOCAL_DATE)
        {// 실제 호출되는 부분.

            string vMessage = string.Empty;

            //초기화//
            mCopy_StartCol = 1;
            mCopy_StartRow = 1;
            mCopy_EndCol = 19;
            mCopy_EndRow = 47;

            mDefaultEndPageRow = 1;
            mDefaultPageRow = 4;    // 페이지 증가후 PageCount 기본값.
            mPrintingLastRow = 45;  //최종 인쇄 라인.
            m1stPrintingLastRow = 45;

            mCurrentRow = 38;

            // 합계.
            mSUM_PL_AMOUNT = 0;
            mSUM_AMOUNT = 0;
            mSUM_GAP_AMOUNT = 0;

            int vTotalRow = 0;
            int vPageRowCount = 0;  //인쇄후 해당 라인 증가 위해.
            int vLIneRow = 0;
            try
            {
                // 실제인쇄되는 행수.
                vTotalRow = pHeader.CurrentRows.Count;

                //mPageTotalNumber = vTotal1ROW / vBy;  // 현재 인쇄 장수 / 총 장수 표시 위해.
                //mPageTotalNumber = (vTotal1ROW % vBy) == 0 ? mPageTotalNumber : (mPageTotalNumber + 1);
                // ? 기준 앞에 비교 문장이고 : 기준으로 앞이 참, 뒤가 거짓.               

                #region ----- Line Write ----

                if (vTotalRow > 0)
                {
                    //실제 라인 row수.
                    vTotalRow = pLine.CurrentRows.Count;

                    HeaderWrite_Etc(pHeader, pApproval_Step, pSOB_DESC, pLOCAL_DATE);

                    // 원본을 복사해서 타깃쉬트에 붙여 넣는다.
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet1, 1);

                    //첫장에 대해서는 시작row부터 체크.
                    vPageRowCount = mCurrentRow - 1;

                    //SetArray1(pGrid, out vGDColumn, out vXLColumn);
                    foreach (System.Data.DataRow vRow in pLine.CurrentRows)
                    {
                        vLIneRow++;
                        vMessage = string.Format("Printing : {0}/{1}", vRow, vTotalRow);
                        mAppInterface.OnAppMessageEvent(vMessage);
                        System.Windows.Forms.Application.DoEvents();

                        mCurrentRow = LineWrite_Etc(vRow, mCurrentRow); // 현재 위치 인쇄 후 다음 인쇄행 리턴.
                        vPageRowCount = vPageRowCount + 1;

                        if (vLIneRow == vTotalRow)
                        {
                            // 마지막 데이터 이면 처리할 사항 기술
                            // 라인지운다 또는 합계를 표시한다 등 기술.
                            SumWrite(mCurrentRow);      //합계.
                        }
                        else
                        {
                            IsNewPage(vPageRowCount);   // 새로운 페이지 체크 및 생성.
                            if (mIsNewPage == true)
                            {
                                mCurrentRow = mCurrentRow + (mCopy_EndRow - vPageRowCount) + mDefaultPageRow;  // 여러장 인쇄시 해당 페이지의 시작되는 위치.
                                vPageRowCount = mDefaultPageRow - 1;
                            }
                        }
                    }
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

        private void IsNewPage(int pPageRowCount)
        {
            if (mPageNumber == 1)
            {
                if (pPageRowCount == m1stPrintingLastRow)
                { // pPrintingLine : 현재 출력된 행.
                    mIsNewPage = true;
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM); 
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
                    mCopyLineSUM = CopyAndPaste(mPrinting, mSourceSheet2, mCopyLineSUM); 
                }
                else
                {
                    mIsNewPage = false;
                }
            }
        }

        #endregion;

        #region ----- Copy&Paste Sheet Method ----

        //[Sheet2]내용을 [Sheet1]에 붙여넣기
        private int CopyAndPaste(XL.XLPrint pPrinting, string pSourceTab, int pCopySumPrintingLine)
        {
            mPageNumber++; //페이지 번호

            int vCopySumPrintingLine = pCopySumPrintingLine;

            mPrinting.XLActiveSheet(pSourceTab); //이 함수를 호출 하지 않으면 그림파일이 XL Sheet에 Insert 되지 않는다.

            //[원본], [Sheet2.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            pPrinting.XLActiveSheet(pSourceTab);
            object vRangeSource = pPrinting.XLGetRange(mCopy_StartRow, mCopy_StartCol, mCopy_EndRow, mCopy_EndCol);

            //[대상], [Sheet1.Cell("A1:AS67")], 엑셀 쉬트에서 복사 시작할 행번호, 엑셀 쉬트에서 복사 시작할 열번호, 엑셀 쉬트에서 복사 종료할 행번호, 엑셀 쉬트에서 복사 종료할 열번호
            int vCopyPrintingRowSTART = pCopySumPrintingLine;

            pPrinting.XLActiveSheet(mTargetSheet);
            object vRangeDestination = pPrinting.XLGetRange(vCopyPrintingRowSTART, mCopy_StartCol, vCopyPrintingRowSTART + mCopy_EndRow, mCopy_EndCol);
            pPrinting.XLCopyRange(vRangeSource, vRangeDestination);

            vCopySumPrintingLine = vCopySumPrintingLine + mCopy_EndRow;
            return vCopySumPrintingLine;
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
            System.IO.DirectoryInfo vWallpaperFolder = new System.IO.DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));

            int vMaxNumber = 1; //MaxIncrement(vWallpaperFolder.ToString(), pSaveFileName);
            vMaxNumber = vMaxNumber + 1;
            string vSaveFileName = string.Format("{0}{1:D3}", pSaveFileName, vMaxNumber);

            vSaveFileName = string.Format("{0}\\{1}.xls", vWallpaperFolder, vSaveFileName);
            mPrinting.XLSave(pSaveFileName);
        }

        #endregion;
    }
}
