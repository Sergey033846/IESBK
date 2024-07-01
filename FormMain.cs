using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

namespace IESBK
{
    public partial class FormMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        // ИЭСБК отчеты - "Отчет о работе контролеров"
        private void barButtonItem30_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Отчет о работе контролеров";
            IWorkbook workbook = form1.spreadsheetControl1.Document;            

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            {
                string otdelenieid = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();
                                
                Worksheet worksheet = workbook.Worksheets[i];
                worksheet.Name = captionotd;
                workbook.Worksheets.Add();

                // задаем смещение табличной части отчета
                int strow = 4;
                int stcol = 0;

                string peryear = "2018";
                string permonth = "11";

                worksheet[0, 0].SetValue("ФЛ без ОДПУ");
                worksheet[1, 0].SetValue(captionotd);
                worksheet[2, 0].SetValue(peryear+", "+permonth);

                worksheet[strow + 0, stcol + 0].SetValue("Вид последнего показания");
                worksheet[strow + 0, stcol + 1].SetValue("Всего л/с");
                worksheet[strow + 0, stcol + 2].SetValue("Всего ПО");
                worksheet[strow + 0, stcol + 3].SetValue("Средний ПО");
                worksheet[strow + 0, stcol + 4].SetValue("Доля л/с, %");
                worksheet[strow + 0, stcol + 5].SetValue("Доля ПО, %");

                string queryString =
                    "SELECT otdelenie.captionotd,pvLPT.propvalue AS typeinfo,COUNT(*) AS totalls, SUM(CAST(REPLACE(pvPOJ.propvalue, ',', '.') AS float)) AS totalpo"+
                    " FROM"+
                    " ([iesbk].[dbo].[tblIESBKlspropvalue] pvLPT"+
                    " LEFT JOIN[iesbk].[dbo].[tblIESBKotdelenie] otdelenie"+
                    " ON otdelenie.otdelenieid = pvLPT.otdelenieid)"+
	                " LEFT JOIN"+
                    " (SELECT pvPO.propvalue, pvPO.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvPO"+
                    " WHERE pvPO.lstypeid = 'ФЛ' AND pvPO.periodyear = '"+peryear+ "' AND pvPO.periodmonth = '" + permonth + "' AND pvPO.lspropertieid = '27' AND pvPO.otdelenieid = '" + otdelenieid+"') pvPOJ"+
                    " ON pvLPT.codeIESBK = pvPOJ.codeIESBK"+
                    " WHERE pvLPT.lstypeid = 'ФЛ' AND pvLPT.periodyear = '" + peryear + "' AND pvLPT.periodmonth = '" + permonth + "' AND pvLPT.lspropertieid = '26' AND pvLPT.otdelenieid = '" + otdelenieid+"'"+

                    " GROUP BY otdelenie.captionotd,pvLPT.propvalue"+
                    " ORDER BY pvLPT.propvalue";
                      
                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                // делаем общий подсчет---------------------------------------
                string queryString2 =
                "SELECT" +
                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '" + peryear + "' AND pv201602.periodmonth = '" + permonth + "' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = '"+otdelenieid+"') AS totalls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '" + peryear + "' AND pv201602.periodmonth = '" + permonth + "' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = '"+otdelenieid +"' AND pv201602.propvalue IS NOT NULL) AS totalpo";

                DataTable tableTOTALlsPO = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALlsPO, dbconnectionStringIESBK, queryString2);

                int totalls = Convert.ToInt32(tableTOTALlsPO.Rows[0]["totalls"]);
                double totalpo = String.IsNullOrWhiteSpace(tableTOTALlsPO.Rows[0]["totalpo"].ToString()) ? 0 :Convert.ToDouble(tableTOTALlsPO.Rows[0]["totalpo"]);
                //------------------------------------------------------------

                for (int j = 0; j < tableTOTALls.Rows.Count; j++)
                {
                    if (String.IsNullOrWhiteSpace(tableTOTALls.Rows[j]["typeinfo"].ToString())) worksheet[strow + j + 1, stcol + 0].SetValue("Расчетное(норматив)");
                    else worksheet[strow + j + 1, stcol + 0].SetValue(tableTOTALls.Rows[j]["typeinfo"].ToString());
                    worksheet[strow + j + 1, stcol + 1].SetValue(tableTOTALls.Rows[j]["totalls"]);
                    worksheet[strow + j + 1, stcol + 2].SetValue(tableTOTALls.Rows[j]["totalpo"]);

                    double totalPO = 0;
                    if (String.IsNullOrWhiteSpace(tableTOTALls.Rows[j]["totalpo"].ToString())) totalPO = 0;
                    else totalPO = Convert.ToDouble(tableTOTALls.Rows[j]["totalpo"]);
                    double vesPO = totalPO / Convert.ToInt32(tableTOTALls.Rows[j]["totalls"]);
                    worksheet[strow + j + 1, stcol + 3].SetValue(vesPO);

                    double dolyaLS = Convert.ToDouble(tableTOTALls.Rows[j]["totalls"]) / totalls*100;
                    worksheet[strow + j + 1, stcol + 4].SetValue(dolyaLS);
                    double dolyaPO = totalPO / totalpo*100;
                    worksheet[strow + j + 1, stcol + 5].SetValue(dolyaPO);

                    worksheet[strow + j + 1, stcol + 1].NumberFormat = "#####";
                    worksheet[strow + j + 1, stcol + 2].NumberFormat = "#";
                    worksheet[strow + j + 1, stcol + 3].NumberFormat = "#.##";
                    worksheet[strow + j + 1, stcol + 4].NumberFormat = "#";
                    worksheet[strow + j + 1, stcol + 5].NumberFormat = "#";
                }

                worksheet.Columns.AutoFit(0, 10);
            } // for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)

            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // ИЭСБК отчеты - "Отчет о работе контролеров"

        // отчет-"шахматка" по наличию л/с и полезного отпуска
        private void barButtonItem34_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
                        
            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Шахматка с января 2018 по ноябрь 2018";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            workbook.History.IsEnabled = false;            
            form1.spreadsheetControl1.BeginUpdate();

            // константы
            //int MAX_PERIOD_MONTH = 12;
            int MAX_PERIOD_MONTH = 11; // январь - ноябрь
            //int MAX_PERIOD_MONTH = 12; // январь - декабрь
            string periodmonthnext = "12"; // переделать

            int columns_in_period_auto = 20+1; // колонок в периоде для автоматического вывода
            int columns_in_period_manual = 4; // колонок в периоде для ручного вывода
            int columns_in_period = columns_in_period_auto + columns_in_period_manual; // общее кол-во колонок в периоде
            int FIRST_COLUMNS = 8;
            int END_COLUMNS = 1 + 3 + 8 + MAX_PERIOD_MONTH + 7;

            //int MAXPROPMAS = 9;
            int MAXCOLinWRKSH = FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + END_COLUMNS;

            /*for (int col = 0; col < MAXCOLinWRKSH; col++)
            {
                worksheet.Columns[col].Font.Name = "Arial";
                worksheet.Columns[col].Font.Size = 8;
            }*/

            DateTime dt_IESBK_MIN = Convert.ToDateTime("01.01.2015"); // левая граница имеющихся данных в OLAP-кубе

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            // загружаем л/с
            DataSetIESBKTableAdapters.tblIESBKlsTableAdapter tblIESBKlsTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlsTableAdapter();
            tblIESBKlsTableAdapter.Fill(DataSetIESBKLoad.tblIESBKls);


            // продумать выборку!!!
            // св-во 36 - "Расход ОДН по нормативу" (убрал)
            /*string queryString = "SELECT DISTINCT codeIESBK,otdelenieid " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +                                
                                "WHERE periodyear = '2016' AND (periodmonth = '01' OR periodmonth = '07') AND lspropertieid='36' AND propvalue IS NULL";*/
            string queryString = "SELECT DISTINCT codeIESBK,otdelenieid " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                 "WHERE periodyear in ('2018') ";
                                 //"AND codeIESBK = 'КНОО00001590'";
                                 //" AND otdelenieid ='" + this.textBox1.Text + "'";
                                
                                //"WHERE periodyear = '2017'";
                                //" AND otdelenieid = 'ЦО'";
                                //+ " AND otdelenieid = '" + this.textBox1.Text + "'";// AND lspropertieid='36' AND propvalue IS NULL";
                                //+ " AND (otdelenieid = 'СОС' OR otdelenieid = 'СОЗ')";// AND lspropertieid='36' AND propvalue IS NULL";
            DataTable tableTOTALls10 = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls10, dbconnectionStringIESBK, queryString);
            //-----------------------------

            // выводим заголовки столбцов -----------------------------------------------------

            // статические
            worksheet[0, 0].SetValue("№ п/п");
            worksheet[0, 1].SetValue("Отделение ИЭСБК");
            worksheet[0, 2].SetValue("Код л/с ИЭСБК");            
            worksheet[0, 3].SetValue("ФИО");                
            worksheet[0, 4].SetValue("Населенный пункт");
            worksheet[0, 5].SetValue("Улица");
            worksheet[0, 6].SetValue("Дом");
            worksheet[0, 7].SetValue("Номер квартиры");
            //worksheet[0, 8].SetValue("Состояние ЛС (на 2016 08)");

            // периодические
            // id "периодических" свойств
            //int[] propidmas = new int[] { 6, 7, 50, 24, 25, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };
            int[] propidmas = new int[] { 51, 20, 24, 25, 6, 7, 50, 26, 27, 28, 55, 29, 30, 53, 31, 54, 32, 33, 34, 35, 36 };
            int idprop_PO_in_propidmas = 7+1; // индекс идентификатора поля ПО от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_lastPOK_in_propidmas = 2+1; // индекс идентификатора поля ПослПоказаниеПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_nomerPU_in_propidmas = 4+1; // индекс идентификатора поля ЗаводскойНомерПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int[] propidmas_doublevalue = new int[] { 27, 28, 55, 29, 30, 53, 31, 54, 32, 33, 34, 35, 36 }; // id числовых полей

            string queryStringlsprop = "SELECT lspropertieid, templateid, numcolumninfile, captionlsprop " +
                                        "FROM [iesbk].[dbo].[tblIESBKlsprop]";
            DataTable tableLSprop = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

            Color Color_IS_PO_NORMATIV_NOT_PU = Color.Orange; // цвет "норматив - безприборник"
            Color Color_IS_PO_NORMATIV_YES_PU = Color.Red; // цвет "норматив - приборник"
            Color Color_IS_PO_SREDNEMES_YES_PU = Color.Blue; // цвет "среднемесячное - приборник"
            Color Color_IS_PO_RASHOD_YES_PU = Color.Green; // цвет "расход по прибору"

            for (int period_i = 1; period_i < MAX_PERIOD_MONTH + 1; period_i++)
            {
                string periodyear = "";
                string periodmonth = "";

                periodyear = "2018";
                periodmonth = (period_i < 10) ? "0"+period_i.ToString() : period_i.ToString();

                /*// сделал для "перехода" года
                if (period_i == 1)
                { 
                    periodyear = "2017";
                    periodmonth = "12";
                }
                else
                if (period_i == 2)
                {
                    periodyear = "2018";
                    periodmonth = "01";
                }*/

                for (int k = 0; k < columns_in_period_auto; k++)
                {
                    DataRow[] lsproprows = tableLSprop.Select("lspropertieid = '" + propidmas[k].ToString() + "'");
                    string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["captionlsprop"].ToString() : null;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(periodyear + " " + periodmonth + ", " + propvaluestr);

                    Color cellColor = Color.Black;
                    // "красим" названия столбцов
                    if (propidmas[k] == 28) cellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (propidmas[k] == 29) cellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (propidmas[k] == 30 || propidmas[k] == 53) cellColor = Color_IS_PO_SREDNEMES_YES_PU;
                    else if (propidmas[k] == 31 || propidmas[k] == 54) cellColor = Color_IS_PO_NORMATIV_YES_PU;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].Font.Color = cellColor;
                }

                tableLSprop.Dispose();

                //------------------------------------------------------------------------------

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(periodyear + " " + periodmonth + ", " + "Расчетный полезный отпуск");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(periodyear + " " + periodmonth + ", " + "Расход по показаниям ПУ");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(periodyear + " " + periodmonth + ", " + "Начисленный полезный отпуск от ИЭСБК");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(periodyear + " " + periodmonth + ", " + "Отклонение (недополученный ПО)");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;
            }

            //------------------------------------------------------------------------------
            //string periodmonthnext = (MAX_PERIOD_MONTH + 1 < 10) ? "0" + (MAX_PERIOD_MONTH + 1).ToString() : (MAX_PERIOD_MONTH + 1).ToString();            
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue("2018" + " " + periodmonthnext + ", " + "среднемесячное (прогноз ПО)");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue("Предыдущее показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue("Предыдущее показание, показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue("Текущее показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue("Текущее показание, показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue("Разница показаний");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue("Разница дней");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue("Среднесуточный расход");

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue("ИТОГО Расход по показаниям ПУ с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].Font.Color = Color.Green;

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue("Начальное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue("Начальное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11].SetValue("Начальное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 12].SetValue("Начальное показание, номер ПУ");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 13].SetValue("Конечное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 14].SetValue("Конечное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 15].SetValue("Конечное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 16].SetValue("Конечное показание, номер ПУ");

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].SetValue("ИТОГО Полезный отпуск от ИЭСБК с начала года (проверка Расхода по показаниям)");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].Font.Color = Color.Blue;

            for (int period_i = 1; period_i <= MAX_PERIOD_MONTH; period_i++)
            {             
                worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17 + period_i].
                    SetValue(worksheet[0, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString());
            }

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].SetValue("ИТОГО Отклонение (недополученный ПО) с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].Font.Color = Color.Red;
            //----------------------------------------------------------

            // главный цикл
            for (int i = 0; i < tableTOTALls10.Rows.Count; i++)
            //for (int i = 0; i < 150; i++)
            {                
                string codeIESBK = tableTOTALls10.Rows[i]["codeIESBK"].ToString();
                //string codeIESBK = "ККОО00019257";

                splashScreenManager1.SetWaitFormDescription(String.Concat(codeIESBK, " ", (i + 1).ToString(), " из ", tableTOTALls10.Rows.Count.ToString(), ")"));

                string otdelenieid = tableTOTALls10.Rows[i]["otdelenieid"].ToString();
                //string otdelenieid = "АО";
                string otdeleniecapt = DataSetIESBKLoad.tblIESBKotdelenie.FindByotdelenieid(otdelenieid)["captionotd"].ToString();
                
                worksheet[i + 1, 0].SetValue((i + 1).ToString());
                worksheet[i + 1, 1].SetValue(otdeleniecapt);
                worksheet[i + 1, 2].SetValue(codeIESBK);
                
                // "красим" строку в зависимости от признака isvalid
                DataRow findrow = DataSetIESBKLoad.tblIESBKls.FindBycodeIESBKlstypeidotdelenieid(codeIESBK, "ФЛ", otdelenieid);
                if (findrow != null)
                {
                    if (findrow["isvalid"].ToString() == "0")
                        worksheet.Rows[i + 1].FillColor = Color.Yellow;                        
                }

                /*// состояние ЛС за июнь 2016  ---------------------------------------
                string periodyear = "2016";
                string periodmonth = "06";                    
                worksheet[i + 1, 3].SetValue(MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 51, SQLconnection));*/

                // основные параметры ЛС за начальный период (01 2017)  ---------------------------------------
                string periodyear = "2018";
                string periodmonth = "01";

                DataRow[] lsproprows;
                /*// "статические" свойства лицевого счета                
                queryStringlsprop = "SELECT codeIESBK, lspropertieid, propvalue " +
                                     "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                     "WHERE codeIESBK='" + codeIESBK + "' AND periodyear = '" + periodyear + "' AND periodmonth = '" + periodmonth + "'";
                tableLSprop = new DataTable();
                MyFUNC_SelectDataFromSQLwoutConnection(tableLSprop, SQLconnection, queryStringlsprop);*/

                // сделано для увеличения производительности
                // фио
                /*string lspropvalue = "";
                                
                lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 5, SQLconnection);
                worksheet[i + 1, 3].SetValue(lspropvalue);

                /*DataRow[] lsproprows = tableLSprop.Select("lspropertieid='5'");
                if (lsproprows.Length > 0) worksheet[i + 1, 3].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // населенный пункт                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 12, SQLconnection);
                worksheet[i + 1, 4].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='12'");
                if (lsproprows.Length > 0) worksheet[i + 1, 4].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // улица                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 13, SQLconnection);
                worksheet[i + 1, 5].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='13'");
                if (lsproprows.Length > 0) worksheet[i + 1, 5].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // дом                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 14, SQLconnection);
                worksheet[i + 1, 6].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='14'");
                if (lsproprows.Length > 0) worksheet[i + 1, 6].SetValue(lsproprows[0]["propvalue"].ToString());                */

                // номер квартиры                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 15, SQLconnection);
                worksheet[i + 1, 7].SetValue(lspropvalue);                
                /*lsproprows = tableLSprop.Select("lspropertieid='15'");
                if (lsproprows.Length > 0) worksheet[i + 1, 7].SetValue(lsproprows[0]["propvalue"].ToString());                */

                //tableLSprop.Dispose();
                //---------------------------------------

                // переменные для подсчета ПО по разнице показаний + сам ПО
                double POKstart = -1;
                double POKend = -1;
                
                DateTime POKstart_date = Convert.ToDateTime("01.01.1900");
                DateTime POKend_date = Convert.ToDateTime("01.01.1900"); 
                string POKstart_kind = null;
                string POKend_kind = null;
                int periodPOKstart = -1; // нумерация с 1 (январь)
                int periodPOKend = -1; // нумерация с 1 (январь)

                string nomerPUstart = null;
                string nomerPUend = null;
                //------------------------------------------------

                // "бежим" по периодическим свойствам лицевого счета
                int START_PERIOD_MONTH = 1;
                for (int period_i = START_PERIOD_MONTH; period_i < START_PERIOD_MONTH + MAX_PERIOD_MONTH; period_i++)
                {
                    periodyear = "2018";
                    periodmonth = (period_i < 10) ? "0"+period_i.ToString() : period_i.ToString();

                    /*// сделал для "перехода" года
                    periodyear = "";
                    periodmonth = "";

                    if (period_i == 1)
                    {
                        periodyear = "2017";
                        periodmonth = "12";
                    }
                    else
                    if (period_i == 2)
                    {
                        periodyear = "2018";
                        periodmonth = "01";
                    }*/

                    //-----------------------------------
                    // "статические" свойства лицевого счета                                

                    string lspropvalue = "";

                    // фио
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 5, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 3].SetValue(lspropvalue);
                    
                    // населенный пункт                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 12, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 4].SetValue(lspropvalue);
                    
                    // улица                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 13, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 5].SetValue(lspropvalue);
                    
                    // дом                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 14, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 6].SetValue(lspropvalue);
                    
                    // номер квартиры                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 15, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 7].SetValue(lspropvalue);
                    //-----------------------------------

                    queryStringlsprop = 
                        String.Concat("SELECT codeIESBK, lspropertieid, propvalue ",
                                     "FROM [iesbk].[dbo].[tblIESBKlspropvalue] ",
                                     "WHERE codeIESBK='", codeIESBK, "' AND periodyear = '", periodyear, "' AND periodmonth = '", periodmonth, "'");
                    tableLSprop = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

                    /*// состояние ЛС (по последнему расчетному периоду)
                    if (period_i == MAX_PERIOD_MONTH)
                    {                           
                        lsproprows = tableLSprop.Select("lspropertieid='51'");                        
                        if (lsproprows.Length > 0) worksheet[i + 1, 8].SetValue(lsproprows[0]["propvalue"].ToString());
                    }*/

                    // флаги для раскраски ячейки полезного отпуска                    
                    bool IS_PO_NORMATIV_NOT_PU = false; // флаг "норматив - безприборник"
                    bool IS_PO_NORMATIV_YES_PU = false; // флаг "норматив - приборник"
                    bool IS_PO_SREDNEMES_YES_PU = false; // флаг "среднемесячное - приборник"
                    bool IS_PO_RASHOD_YES_PU = false; // флаг "расход по прибору"

                    // выводим "периодические" поля - сделал в цикле                    
                    for (int k = 0; k < columns_in_period_auto; k++)
                    {
                        /*string strTEST = "12345678";
                        string propvaluestr = strTEST;*/

                        lsproprows = tableLSprop.Select(String.Concat("lspropertieid = '", propidmas[k].ToString(), "'"));

                        string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;                        
                        double? propvalue = null;

                        if (System.Array.IndexOf(propidmas_doublevalue, propidmas[k]) >= 0)
                        {
                            if (!String.IsNullOrWhiteSpace(propvaluestr)) propvalue = Convert.ToDouble(propvaluestr);
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvalue);
                        }
                        else
                        {
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvaluestr);
                        }

                        // формируем флаги для раскраски ячейки полезного отпуска
                        // помним о массиве периодических свойств
                        //int[] propidmas = new int[] { 6, 7, 50, 24, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };                        
                        if (propvalue != null && propvalue != 0)
                        {
                            if (propidmas[k] == 28) IS_PO_NORMATIV_NOT_PU = true;
                            else if (propidmas[k] == 29) IS_PO_RASHOD_YES_PU = true;
                            else if (propidmas[k] == 30) IS_PO_SREDNEMES_YES_PU = true;
                            else if (propidmas[k] == 31) IS_PO_NORMATIV_YES_PU = true;
                        };
                        //-------------------------------------------------------
                    } // выводим "периодические" поля - сделал в цикле

                    // раскрашиваем колонку ПолОтп в зависимости от "слагаемых"
                    // помним о массиве периодических свойств
                    //int[] propidmas = new int[] { 6, 7, 50, 24, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };

                    Color PolOtpCellColor = Color.Black;
                    if (IS_PO_RASHOD_YES_PU) PolOtpCellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (IS_PO_NORMATIV_NOT_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (IS_PO_NORMATIV_YES_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_YES_PU;
                    else if (IS_PO_SREDNEMES_YES_PU) PolOtpCellColor = Color_IS_PO_SREDNEMES_YES_PU;                    
                    worksheet[i + 1, FIRST_COLUMNS + (period_i - 1 ) * columns_in_period + idprop_PO_in_propidmas].Font.Color = PolOtpCellColor; // "красим" ПолезныйОтпуск - propid = 27

                    //-----------------------------------------------------------------
                    // формируем "периодические" колонки "Расход по показаниям" и "Отклонение (недополученный ПО)", если не было замены ПУ

                    if (period_i > START_PERIOD_MONTH) // пропускаем первый месяц, т.к. в нем не найдем "предыдущих показаний"
                    {
                        string propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        double POKend_period = -1;
                        //double? POIESBK_period = null;
                        double POIESBK_period = 0;
                        double POIESBKend_period = 0; // ПО последнего периода

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";"))
                        {
                            POKend_period = Convert.ToDouble(propvaluestr);

                            double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                            POIESBK_period += POcellvalue;

                            POIESBKend_period = POcellvalue;
                        }

                        string nomerPUend_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        /*lsproprows = tableLSprop.Select("lspropertieid='22'"); // предыдущее показание ПУ
                        propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                        int period_i_step = 1;
                        double POKstart_period = -1;
                        string nomerPUstart_period = null;
                        propvaluestr = null;
                                                
                        do
                        {                            
                            propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();                         
                            nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();

                            // суммируем полезный отпуск, пропуская начальный интервал
                            if (String.IsNullOrWhiteSpace(propvaluestr))
                            {
                                double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.NumericValue;
                                POIESBK_period += POcellvalue;
                            }

                            period_i_step += 1;

                        } while (String.IsNullOrWhiteSpace(propvaluestr) && period_i_step < period_i);

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POKstart_period = Convert.ToDouble(propvaluestr);

                        //nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 2) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        // если имеются оба показания и не было замены ПУ, то считаем значения
                        if (POKend_period != -1 && POKstart_period != -1 && String.Equals(nomerPUstart_period, nomerPUend_period))
                        //if (POKend_period != -1 && POKstart_period != -1 && POKstart_period <= POKend_period)                         
                        {
                            /*propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                            double? POIESBK_period = null;
                            if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POIESBK_period = Convert.ToDouble(propvaluestr);*/

                            double POIESBKPU_period = POKend_period - POKstart_period;
                            double POIESBKDelta_period = POIESBKPU_period - POIESBK_period;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(POIESBKend_period + POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(POIESBKPU_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(POIESBK_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;                            
                        }

                    } // if (period_i > START_PERIOD_MONTH) 
                    //-----------------------------------------------------------------

                    // ищем стартовое и конечное показание для расчета ИТОГОВОГО ПО по показаниям (за все периоды)                  
                    lsproprows = tableLSprop.Select("lspropertieid='25'");
                    string pokstr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                    
                    if (!String.IsNullOrWhiteSpace(pokstr) && !pokstr.Contains(";"))
                    {
                        if (POKstart == -1)
                        {
                            POKstart = Convert.ToDouble(pokstr);
                            periodPOKstart = period_i;

                            lsproprows = tableLSprop.Select("lspropertieid='7'");
                            nomerPUstart = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;

                            lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKstart_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());

                            lsproprows = tableLSprop.Select("lspropertieid='26'");
                            POKstart_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                        }
                        else
                        {
                            // поменять местами нижнее условие и присовение значение конечному показанию!!!!!
                            POKend = Convert.ToDouble(pokstr);
                            periodPOKend = period_i;

                            lsproprows = tableLSprop.Select("lspropertieid='7'");
                            nomerPUend = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;

                            lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKend_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());

                            lsproprows = tableLSprop.Select("lspropertieid='26'");
                            POKend_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;

                            if (!String.Equals(nomerPUstart, nomerPUend)) // если номера ПУ не равны, то последнее делаем начальным
                            {
                                nomerPUstart = nomerPUend;
                                POKstart = POKend;
                                POKstart_date = POKend_date;
                                POKstart_kind = POKend_kind;

                                POKend = -1;

                                periodPOKstart = periodPOKend;
                                periodPOKend = -1;                        
                            };
                        }                        
                    }
                    
                    //-----------------------------------------------------------------

                    tableLSprop.Dispose();

                } // for (int period_i = 1; period_i < 7; period_i++)

                if (POKstart != -1 && POKend != -1) // если имеются оба показания ПУ (для расчета ПО по показаниям ПУ)
                {
                    /*// суммируем полезный отпуск от ИЭСБК со следующего расчетного периода, которому предшествовало показание ПУ
                    lsproprows = tableLSprop.Select("lspropertieid='27'");
                    string postr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                    if (POKstart != -1 && POKend == -1 && !String.IsNullOrWhiteSpace(postr)) POIESBKTotal += Convert.ToDouble(postr);*/

                    // суммируем полезный отпуск от ИЭСБК по ранее заполненным колонкам со следующего расчетного периода, которому предшествовало показание ПУ
                    double POIESBKTotal = 0;
                    for (int period_i = periodPOKstart+1; period_i <= periodPOKend; period_i++)
                    {
                        double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                        POIESBKTotal += POcellvalue;

                        // выводим расшифровку формирования ПО
                        worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17 + period_i].SetValue(POcellvalue);
                    }

                    double POIESBKPU = POKend - POKstart;
                    double POIESBKDelta = POIESBKPU - POIESBKTotal;

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue(POIESBKPU);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].Font.Color = Color.Green;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue(POKstart_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue(POKstart);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11].SetValue(POKstart_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 12].SetValue(nomerPUstart);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 13].SetValue(POKend_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 14].SetValue(POKend);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 15].SetValue(POKend_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 16].SetValue(nomerPUend);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].SetValue(POIESBKTotal);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].Font.Color = Color.Blue;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].SetValue(POIESBKDelta);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].Font.Color = Color.Red;

                    // beg расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)

                    string periodyeartek = "2018"; // было 2017
                    //string periodmonthtek = (MAX_PERIOD_MONTH < 10) ? "0" + MAX_PERIOD_MONTH.ToString() : MAX_PERIOD_MONTH.ToString();

                    // сделал для перехода года, потом убрать!!!
                    string periodmonthtek = (MAX_PERIOD_MONTH-1 < 10) ? "0" + (MAX_PERIOD_MONTH-1).ToString() : (MAX_PERIOD_MONTH-1).ToString();
                    string codels = codeIESBK;

                    // ищем ближайшее "правое" показание                
                    string value_right = null;
                    string dtvalue_right = null;

                    DateTime dt_right = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);
                    string year_right = null;
                    string month_right = null;

                    dt_right = dt_right.AddMonths(+1); // учитываем текущий месяц, т.е. +1-1 = 0

                    while (String.IsNullOrWhiteSpace(value_right) && dt_right >= dt_IESBK_MIN)
                    {
                        dt_right = dt_right.AddMonths(-1);
                        year_right = dt_right.Year.ToString();
                        month_right = null;
                        if (dt_right.Month < 10) month_right = "0" + dt_right.Month.ToString();
                        else month_right = dt_right.Month.ToString();

                        value_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 25, SQLconnection); // свойство "Текущее показание ПУ"                                        
                    };
                    //----------------------------------

                    // ищем "левое" показание, при условии, что нашли "правое" ------------------
                    string value_left = null;
                    string dtvalue_left = null;
                    string year_left = null;
                    string month_left = null;

                    if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    {
                        // получаем дату "правого" показания
                        dtvalue_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                        /*// выводим информацию о "правом" показании
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(dtvalue_right); // дата

                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(value_right); // показание

                        string rightpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_right, month_right, 26, SQLconnection);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(rightpok_type); // вид*/

                        DateTime dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-5); // было -6, отматываем -5-1 = -6 мес. = 180 дней от "правого" показания
                        
                        while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN)
                        {
                            dt_left = dt_left.AddMonths(-1);
                            year_left = dt_left.Year.ToString();
                            month_left = null;
                            if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                            else month_left = dt_left.Month.ToString();

                            value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                        };

                        // если нет данных да период не менее 6 мес., то ищем за в периоде [6 мес.;3 мес.]                        
                        if (String.IsNullOrWhiteSpace(value_left))
                        {
                            dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-7); // отматываем 7 мес., т.к. в теле цикла сразу +1, т.е. -7+1 = -6
                            
                            DateTime dt_IESBK_left_MAX = Convert.ToDateTime(dtvalue_right).AddMonths(-3);

                            while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN && dt_left < dt_IESBK_left_MAX)
                            {
                                dt_left = dt_left.AddMonths(+1);
                                year_left = dt_left.Year.ToString();
                                month_left = null;
                                if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                                else month_left = dt_left.Month.ToString();

                                value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                            };
                        } // if (String.IsNullOrWhiteSpace(value_left)) // если нет данных за период не менее 6 мес.

                        // получаем дату "левого" показания и выводим информацию о нем
                        if (!String.IsNullOrWhiteSpace(value_left))
                        {
                            dtvalue_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                            /*// выводим информацию о "правом" показании
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(dtvalue_left); // дата

                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(value_left); // показание

                            string leftpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_left, month_left, 26, SQLconnection);
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(leftpok_type); // вид*/
                        }

                        // если даты "левого" и "правого" показаний не пустые, то формируем расчет среднемесячного
                        if (!String.IsNullOrWhiteSpace(dtvalue_left) && !dtvalue_left.Contains(";") && !value_left.Contains(";") && !String.IsNullOrWhiteSpace(dtvalue_right))
                        {                            
                            double pokleft = Convert.ToDouble(value_left);
                            double pokright = Convert.ToDouble(value_right);

                            // если не нарушен нарастающий итог
                            if (pokleft <= pokright)
                            {
                                System.TimeSpan deltaday = Convert.ToDateTime(dtvalue_right) - Convert.ToDateTime(dtvalue_left);
                                double deltapok = pokright - pokleft;

                                double srednesut_calc = deltapok / deltaday.Days;
                                double srmes_calc = Math.Round(srednesut_calc * DateTime.DaysInMonth(Convert.ToInt32(periodyeartek), Convert.ToInt32(periodmonthtek)));

                                // формируем отчет -----------------------------------------------

                                /*if (!String.IsNullOrWhiteSpace(srmes_iesbk_str)) // еслм СрМес ПО ИЭСБК присутствует
                                {
                                    DateTime dt_period = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);

                                    if (srmes_iesbk >= 0 && Convert.ToDateTime(dtvalue_right).CompareTo(dt_period) < 0) // не выводим наши расчеты, если СрМес ИЭСБК < 0 и правое показание принадлежит текущему периоду анализа
                                    {
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(srmes_calc); // СрМес РАСЧ                                

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Red;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(srmes_calc - srmes_iesbk); // Недополученный ПО                                
                                    }
                                }*/

                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue(srmes_calc);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;

                                // выводим расшифровку расчета Прогноза (СрМес)
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue(dtvalue_left);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue(pokleft);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue(dtvalue_right);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue(pokright);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue(deltapok);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue(deltaday.Days);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue(srednesut_calc);

                            } // if (pokleft <= pokright)
                              //-----------------------------------------------------------------
                        } // if (!String.IsNullOrWhiteSpace(dtvalue_left) && !String.IsNullOrWhiteSpace(dtvalue_right))

                    } // if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    // end расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)
                }

                //splashScreenManager1.SetWaitFormDescription(String.Concat("Обработка данных (", (i + 1).ToString(), " из ", tableTOTALls10.Rows.Count.ToString(), ")"));
                //splashScreenManager1.SetWaitFormDescription(String.Concat(codeIESBK, " ", (i + 1).ToString(), " из ", tableTOTALls10.Rows.Count.ToString(), ")"));

            } // for (int i = 0; i < tableTOTALls10.Rows.Count; i++)

            for (int col = 0; col < MAXCOLinWRKSH; col++)
            {
                worksheet.Columns[col].Font.Name = "Arial";
                worksheet.Columns[col].Font.Size = 8;
            }

            // форматируем строку-заголовок
            worksheet.Rows[0].Font.Bold = true;
            worksheet.Rows[0].Alignment.WrapText = true;
            worksheet.Rows[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Rows[0].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            worksheet.Rows[0].AutoFit();

            worksheet.Columns.AutoFit(0, MAXCOLinWRKSH);

            worksheet.Columns.Group(3, 7, true); // группируем по колонкам "ФИО" - "Номер квартиры"            

            // группируем колонки расшифровки прогноза СрМес
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7,true);
            
            //int[] propidmas = new int[] { 6, 7, 50, 24, 25, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };
            //int[] propidmas = new int[] { 51, 24, 25, 6, 7, 50, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };

            // группируем "периодические" значения - в частности расшифровку ПО
            // КРИВО!!!! ПРИ ДОБАВЛЕНИИ СВОЙСТВА ВПЕРИОДИЧЕСКУЮ СЕКЦИЮ ЗДЕСЬ ПРИХОДИТСЯ ДОБАВЛЯТЬ!!!
            for (int period_i = 0; period_i < MAX_PERIOD_MONTH; period_i++)
            {
                //worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period, FIRST_COLUMNS + period_i * columns_in_period + 2, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 2 + 1, FIRST_COLUMNS + period_i * columns_in_period + 4 + 1 + 1, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 7 + 1 + 1, FIRST_COLUMNS + period_i * columns_in_period + 7 + 10 + 2 + 1, true); // после "ПолОтп"
            }

            // группируем последние колонки анализа ПО по показаниям ПУ
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1 + 1 + 7, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1 + 7, true);

            // группируем колонки слагаемых ПО от ИЭСБК
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 2 + 1 + 7, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1 + MAX_PERIOD_MONTH + 1 + 7, true);                                    

            worksheet.FreezeRows(0); // "фиксируем" верхнюю строку

            form1.spreadsheetControl1.EndUpdate();

            SQLconnection.Close();
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // отчет-"шахматка" по наличию л/с и полезного отпуска

        // ИЭСБК отчеты - "Кол-во л/с и ПО"
        private void barButtonItem35_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "ИЭСБК кол-во лс и полезный отпуск";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            string queryString =
                "SELECT otd.captionotd," +
                    " (SELECT COUNT(*)"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602"+
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '01' AND pv201602.lspropertieid = '3' AND"+
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls,"+
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602"+
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '01' AND pv201602.lspropertieid = '27' AND"+
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po,"+

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '02' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '02' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '03' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '03' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '04' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '04' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '05' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '05' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '06' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '06' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '07' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '07' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '08' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '08' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '09' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '09' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '10' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '10' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '11' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '11' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po " +

                    /*" (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2017' AND pv201602.periodmonth = '12' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2017' AND pv201602.periodmonth = '12' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po" +*/

                " FROM[iesbk].[dbo].tblIESBKotdelenie otd" +
                " GROUP BY otd.otdelenieid,otd.captionotd" +
                " ORDER BY otd.otdelenieid";

            DataTable tableTOTALls10 = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls10, dbconnectionStringIESBK, queryString);
            //-----------------------------

            // задаем смещение табличной части отчета
            int strow = 2;
            int stcol = 0;

            worksheet[0, 0].SetValue("ФЛ без ОДПУ");

            worksheet[strow + 0, stcol + 0].SetValue("Отделение ИЭСБК");            
            worksheet[strow + 0, stcol + 1].SetValue("2018 01 лс");
            worksheet[strow + 0, stcol + 2].SetValue("2018 01 ПО");
            worksheet[strow + 0, stcol + 3].SetValue("2018 01 среднее");
            worksheet[strow + 0, stcol + 4].SetValue("2018 02 лс");
            worksheet[strow + 0, stcol + 5].SetValue("2018 02 ПО");
            worksheet[strow + 0, stcol + 6].SetValue("2018 02 среднее");
            worksheet[strow + 0, stcol + 7].SetValue("2018 03 лс");
            worksheet[strow + 0, stcol + 8].SetValue("2018 03 ПО");
            worksheet[strow + 0, stcol + 9].SetValue("2018 03 среднее");
            worksheet[strow + 0, stcol + 10].SetValue("2018 04 лс");
            worksheet[strow + 0, stcol + 11].SetValue("2018 04 ПО");
            worksheet[strow + 0, stcol + 12].SetValue("2018 04 среднее");
            worksheet[strow + 0, stcol + 13].SetValue("2018 05 лс");
            worksheet[strow + 0, stcol + 14].SetValue("2018 05 ПО");
            worksheet[strow + 0, stcol + 15].SetValue("2018 05 среднее");
            worksheet[strow + 0, stcol + 16].SetValue("2018 06 лс");
            worksheet[strow + 0, stcol + 17].SetValue("2018 06 ПО");
            worksheet[strow + 0, stcol + 18].SetValue("2018 06 среднее");
            worksheet[strow + 0, stcol + 19].SetValue("2018 07 лс");
            worksheet[strow + 0, stcol + 20].SetValue("2018 07 ПО");
            worksheet[strow + 0, stcol + 21].SetValue("2018 07 среднее");
            worksheet[strow + 0, stcol + 22].SetValue("2018 08 лс");
            worksheet[strow + 0, stcol + 23].SetValue("2018 08 ПО");
            worksheet[strow + 0, stcol + 24].SetValue("2018 08 среднее");
            worksheet[strow + 0, stcol + 25].SetValue("2018 09 лс");
            worksheet[strow + 0, stcol + 26].SetValue("2018 09 ПО");
            worksheet[strow + 0, stcol + 27].SetValue("2018 09 среднее");
            worksheet[strow + 0, stcol + 28].SetValue("2018 10 лс");
            worksheet[strow + 0, stcol + 29].SetValue("2018 10 ПО");
            worksheet[strow + 0, stcol + 30].SetValue("2018 10 среднее");
            worksheet[strow + 0, stcol + 31].SetValue("2018 11 лс");
            worksheet[strow + 0, stcol + 32].SetValue("2018 11 ПО");
            worksheet[strow + 0, stcol + 33].SetValue("2018 11 среднее");
            /*worksheet[strow + 0, stcol + 34].SetValue("2017 12 лс");
            worksheet[strow + 0, stcol + 35].SetValue("2017 12 ПО");
            worksheet[strow + 0, stcol + 36].SetValue("2017 12 среднее");*/

            for (int i = 0; i < tableTOTALls10.Rows.Count; i++)            
            {                   
                worksheet[strow + i + 1, stcol + 0].SetValue(tableTOTALls10.Rows[i][0].ToString());

                double? dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][2].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][1].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][2].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][1].ToString());
                    worksheet[strow + i + 1, stcol + 1].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][1].ToString()));
                    worksheet[strow + i + 1, stcol + 2].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][2].ToString()));
                    worksheet[strow + i + 1, stcol + 3].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 1].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 2].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 3].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][4].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][3].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][4].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][3].ToString());
                    worksheet[strow + i + 1, stcol + 4].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][3].ToString()));
                    worksheet[strow + i + 1, stcol + 5].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][4].ToString()));
                    worksheet[strow + i + 1, stcol + 6].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 4].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 5].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 6].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][6].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][5].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][6].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][5].ToString());
                    worksheet[strow + i + 1, stcol + 7].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][5].ToString()));
                    worksheet[strow + i + 1, stcol + 8].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][6].ToString()));
                    worksheet[strow + i + 1, stcol + 9].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 7].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 8].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 9].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][8].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][7].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][8].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][7].ToString());
                    worksheet[strow + i + 1, stcol + 10].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][7].ToString()));
                    worksheet[strow + i + 1, stcol + 11].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][8].ToString()));
                    worksheet[strow + i + 1, stcol + 12].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 10].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 11].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 12].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][10].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][9].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][10].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][9].ToString());
                    worksheet[strow + i + 1, stcol + 13].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][9].ToString()));
                    worksheet[strow + i + 1, stcol + 14].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][10].ToString()));
                    worksheet[strow + i + 1, stcol + 15].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 13].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 14].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 15].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][12].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][11].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][12].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][11].ToString());
                    worksheet[strow + i + 1, stcol + 16].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][11].ToString()));
                    worksheet[strow + i + 1, stcol + 17].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][12].ToString()));
                    worksheet[strow + i + 1, stcol + 18].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 16].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 17].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 18].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][14].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][13].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][14].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][13].ToString());
                    worksheet[strow + i + 1, stcol + 19].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][13].ToString()));
                    worksheet[strow + i + 1, stcol + 20].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][14].ToString()));
                    worksheet[strow + i + 1, stcol + 21].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 19].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 20].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 21].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][16].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][15].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][16].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][15].ToString());
                    worksheet[strow + i + 1, stcol + 22].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][15].ToString()));
                    worksheet[strow + i + 1, stcol + 23].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][16].ToString()));
                    worksheet[strow + i + 1, stcol + 24].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 22].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 23].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 24].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][18].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][17].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][18].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][17].ToString());
                    worksheet[strow + i + 1, stcol + 25].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][17].ToString()));
                    worksheet[strow + i + 1, stcol + 26].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][18].ToString()));
                    worksheet[strow + i + 1, stcol + 27].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 25].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 26].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 27].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][20].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][19].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][20].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][19].ToString());
                    worksheet[strow + i + 1, stcol + 28].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][19].ToString()));
                    worksheet[strow + i + 1, stcol + 29].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][20].ToString()));
                    worksheet[strow + i + 1, stcol + 30].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 28].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 29].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 30].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][22].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][21].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][22].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][21].ToString());
                    worksheet[strow + i + 1, stcol + 31].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][21].ToString()));
                    worksheet[strow + i + 1, stcol + 32].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][22].ToString()));
                    worksheet[strow + i + 1, stcol + 33].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 31].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 32].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 33].NumberFormat = "#.##";

                /*//dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][24].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][23].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][24].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][23].ToString());
                    worksheet[strow + i + 1, stcol + 34].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][23].ToString()));
                    worksheet[strow + i + 1, stcol + 35].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][24].ToString()));
                    worksheet[strow + i + 1, stcol + 36].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 34].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 35].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 36].NumberFormat = "#.##";*/

                splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + ")");
            } // for (int i = 0; i < tableTOTALls10.Rows.Count; i++)

            worksheet.Columns.AutoFit(0, 49); // переделать!!!
            worksheet.FreezeColumns(0);

            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // ИЭСБК отчеты - "Кол-во л/с и ПО"

        // ИЭСБК отчеты - "Нарушение нарастающего итога"
        private void barButtonItem36_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Нарушение нарастающего итога показаний";
            IWorkbook workbook = form1.spreadsheetControl1.Document;

            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            int MAXCOLinWRKSH = 17;
            Color infocolor_prev = Color.DimGray;

            //for (int i = 0; i < 1; i++)
            for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            {
                string otdelenieid = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();

                Worksheet worksheet = workbook.Worksheets[i];
                worksheet.Name = captionotd;
                workbook.Worksheets.Add();

                // задаем смещение табличной части отчета
                int strow = 6;
                int stcol = 0;

                string peryearprev = "2018";
                string permonthprev = "10";

                string peryear = "2018";
                string permonth = "11";

                for (int col = 0; col < MAXCOLinWRKSH; col++)
                {
                    worksheet.Columns[col].Font.Name = "Arial";
                    worksheet.Columns[col].Font.Size = 8;
                }

                worksheet[0, 0].SetValue("Нарушение нарастающего итога");
                worksheet[0, 0].Font.Bold = true;

                worksheet[1, 0].SetValue("3.3.2 ФЛ без ОДПУ");
                worksheet[2, 0].SetValue(captionotd);
                worksheet[3, 0].SetValue("Текущий период:" );
                worksheet[3, 1].SetValue(peryear + ", " + permonth);
                worksheet[4, 0].SetValue("Предыдущий период:");
                worksheet[4, 1].SetValue(peryearprev + ", " + permonthprev);

                worksheet[strow + 0, stcol + 0].SetValue("Код ИЭСБК");
                worksheet[strow + 0, stcol + 1].SetValue("ФИО");

                worksheet[strow + 0, stcol + 2].SetValue("Район");
                worksheet[strow + 0, stcol + 3].SetValue("Населенный пункт");
                worksheet[strow + 0, stcol + 4].SetValue("Улица");
                worksheet[strow + 0, stcol + 5].SetValue("Дом");
                worksheet[strow + 0, stcol + 6].SetValue("Номер квартиры");

                string periodstr = peryear + " " + permonth + ", ";
                string periodstrprev = peryearprev + " " + permonthprev + ", ";
                worksheet[strow + 0, stcol + 7].SetValue(periodstrprev + "предыдущее показание, дата");
                worksheet[strow + 0, stcol + 7].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 8].SetValue(periodstrprev + "предыдущее показание, значение");
                worksheet[strow + 0, stcol + 8].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 9].SetValue(periodstrprev + "предыдущее показание, номер ПУ");
                worksheet[strow + 0, stcol + 9].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 10].SetValue(periodstrprev + "предыдущее показание, источник");
                worksheet[strow + 0, stcol + 10].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 11].SetValue(periodstr + "начальное показание, дата");
                worksheet[strow + 0, stcol + 12].SetValue(periodstr + "начальное показание, значение");
                worksheet[strow + 0, stcol + 13].SetValue(periodstr + "начальное показание, номер ПУ");
                worksheet[strow + 0, stcol + 14].SetValue(periodstr + "начальное показание, источник");

                worksheet[strow + 0, stcol + 15].SetValue("НАРУШЕНИЕ значения при одинаковом номере ПУ");
                worksheet[strow + 0, stcol + 15].Font.Color = Color.Red;
                worksheet[strow + 0, stcol + 16].SetValue("НАРУШЕНИЕ даты при одинаковом номере ПУ");
                worksheet[strow + 0, stcol + 16].Font.Color = Color.Blue;
                //-----------------------

                string queryString =
                    "SELECT otdelenie.captionotd,"+
	                "pvStartPok2.codeIESBK AS codels,"+

                    "pvStartPokFIO.propvalue AS pvStartPokFIOvalue, " +
                    "pvStartPokAddr1.propvalue AS pvStartPokAddr1value, pvStartPokAddr2.propvalue AS pvStartPokAddr2value," +
                    "pvStartPokAddr3.propvalue AS pvStartPokAddr3value, pvStartPokAddr4.propvalue AS pvStartPokAddr4value," +
                    "pvStartPokAddr5.propvalue AS pvStartPokAddr5value, " +

                    "pvStartPokPUNomer.propvalue AS StartPokPUNomervalue, " +
                    "pvEndPokPUNomer.propvalue AS prevPokPUNomervalue, " +

                    "pvEndPok1.propvalue AS prevPOKdatePrevPer,pvEndPok2.propvalue AS prevPOKvaluePrevPer, pvEndPok3.propvalue AS prevPOKtypePrevPer," +
	                "pvStartPok1.propvalue AS startPOKdateTekPer,pvStartPok2.propvalue AS startPOKvalueTekPer,pvStartPok3.propvalue AS startPOKtypeTekPer"+
                    " FROM"+
                    " ([iesbk].[dbo].[tblIESBKlspropvalue] pvStartPok2"+
                    " LEFT JOIN[iesbk].[dbo].[tblIESBKotdelenie] otdelenie"+
                    " ON otdelenie.otdelenieid = pvStartPok2.otdelenieid)"+
	                " LEFT JOIN"+
                    " (SELECT pvStartPok1.propvalue, pvStartPok1.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPok1"+
                    " WHERE pvStartPok1.lstypeid = 'ФЛ' AND pvStartPok1.periodyear = '"+peryear+ "' AND pvStartPok1.periodmonth = '" + permonth + "' AND pvStartPok1.lspropertieid = '21' AND pvStartPok1.otdelenieid = '" + otdelenieid + "') pvStartPok1" +
                    " ON pvStartPok2.codeIESBK = pvStartPok1.codeIESBK"+
                    " LEFT JOIN"+
                    " (SELECT pvStartPok3.propvalue, pvStartPok3.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPok3"+
                    " WHERE pvStartPok3.lstypeid = 'ФЛ' AND pvStartPok3.periodyear = '" + peryear + "' AND pvStartPok3.periodmonth = '" + permonth + "' AND pvStartPok3.lspropertieid = '23' AND pvStartPok3.otdelenieid = '" + otdelenieid + "') pvStartPok3" +
                    " ON pvStartPok2.codeIESBK = pvStartPok3.codeIESBK"+

                    // добавляем поля ФИО и адреса
                    " LEFT JOIN" +
                    " (SELECT pvStartPokFIO.propvalue, pvStartPokFIO.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokFIO" +
                    " WHERE pvStartPokFIO.lstypeid = 'ФЛ' AND pvStartPokFIO.periodyear = '" + peryear + "' AND pvStartPokFIO.periodmonth = '" + permonth + "' AND pvStartPokFIO.lspropertieid = '5' AND pvStartPokFIO.otdelenieid = '" + otdelenieid + "') pvStartPokFIO" +
                    " ON pvStartPok2.codeIESBK = pvStartPokFIO.codeIESBK" +

                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr1.propvalue, pvStartPokAddr1.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr1" +
                    " WHERE pvStartPokAddr1.lstypeid = 'ФЛ' AND pvStartPokAddr1.periodyear = '" + peryear + "' AND pvStartPokAddr1.periodmonth = '" + permonth + "' AND pvStartPokAddr1.lspropertieid = '11' AND pvStartPokAddr1.otdelenieid = '" + otdelenieid + "') pvStartPokAddr1" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr1.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr2.propvalue, pvStartPokAddr2.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr2" +
                    " WHERE pvStartPokAddr2.lstypeid = 'ФЛ' AND pvStartPokAddr2.periodyear = '" + peryear + "' AND pvStartPokAddr2.periodmonth = '" + permonth + "' AND pvStartPokAddr2.lspropertieid = '12' AND pvStartPokAddr2.otdelenieid = '" + otdelenieid + "') pvStartPokAddr2" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr2.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr3.propvalue, pvStartPokAddr3.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr3" +
                    " WHERE pvStartPokAddr3.lstypeid = 'ФЛ' AND pvStartPokAddr3.periodyear = '" + peryear + "' AND pvStartPokAddr3.periodmonth = '" + permonth + "' AND pvStartPokAddr3.lspropertieid = '13' AND pvStartPokAddr3.otdelenieid = '" + otdelenieid + "') pvStartPokAddr3" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr3.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr4.propvalue, pvStartPokAddr4.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr4" +
                    " WHERE pvStartPokAddr4.lstypeid = 'ФЛ' AND pvStartPokAddr4.periodyear = '" + peryear + "' AND pvStartPokAddr4.periodmonth = '" + permonth + "' AND pvStartPokAddr4.lspropertieid = '14' AND pvStartPokAddr4.otdelenieid = '" + otdelenieid + "') pvStartPokAddr4" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr4.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr5.propvalue, pvStartPokAddr5.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr5" +
                    " WHERE pvStartPokAddr5.lstypeid = 'ФЛ' AND pvStartPokAddr5.periodyear = '" + peryear + "' AND pvStartPokAddr5.periodmonth = '" + permonth + "' AND pvStartPokAddr5.lspropertieid = '15' AND pvStartPokAddr5.otdelenieid = '" + otdelenieid + "') pvStartPokAddr5" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr5.codeIESBK" +
                    //----------------------

                    // номер текущего ПУ
                    " LEFT JOIN" +
                    " (SELECT pvStartPokPUNomer.propvalue, pvStartPokPUNomer.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokPUNomer" +
                    " WHERE pvStartPokPUNomer.lstypeid = 'ФЛ' AND pvStartPokPUNomer.periodyear = '" + peryear + "' AND pvStartPokPUNomer.periodmonth = '" + permonth + "' AND pvStartPokPUNomer.lspropertieid = '7' AND pvStartPokPUNomer.otdelenieid = '" + otdelenieid + "') pvStartPokPUNomer" +
                    " ON pvStartPok2.codeIESBK = pvStartPokPUNomer.codeIESBK" +
                    //----------------------

                    " LEFT JOIN" +
                    " (SELECT pvEndPok2.propvalue, pvEndPok2.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPok2"+
                    " WHERE pvEndPok2.lstypeid = 'ФЛ' AND pvEndPok2.periodyear = '" + peryearprev + "' AND pvEndPok2.periodmonth = '" + permonthprev + "' AND pvEndPok2.lspropertieid = '25' AND pvEndPok2.otdelenieid = '" + otdelenieid + "') pvEndPok2" +
                    " ON pvStartPok2.codeIESBK = pvEndPok2.codeIESBK"+
                    " LEFT JOIN" +
                    " (SELECT pvEndPok1.propvalue, pvEndPok1.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPok1"+
                    " WHERE pvEndPok1.lstypeid = 'ФЛ' AND pvEndPok1.periodyear = '" + peryearprev + "' AND pvEndPok1.periodmonth = '" + permonthprev + "' AND pvEndPok1.lspropertieid = '24' AND pvEndPok1.otdelenieid = '" + otdelenieid + "') pvEndPok1" +
                    " ON pvStartPok2.codeIESBK = pvEndPok1.codeIESBK"+
                    " LEFT JOIN" +
                    " (SELECT pvEndPok3.propvalue, pvEndPok3.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPok3"+
                    " WHERE pvEndPok3.lstypeid = 'ФЛ' AND pvEndPok3.periodyear = '" + peryearprev + "' AND pvEndPok3.periodmonth = '" + permonthprev + "' AND pvEndPok3.lspropertieid = '26' AND pvEndPok3.otdelenieid = '" + otdelenieid + "') pvEndPok3" +
                    " ON pvStartPok2.codeIESBK = pvEndPok3.codeIESBK"+

                    // номер предыдущего ПУ
                    " LEFT JOIN" +
                    " (SELECT pvEndPokPUNomer.propvalue, pvEndPokPUNomer.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPokPUNomer" +
                    " WHERE pvEndPokPUNomer.lstypeid = 'ФЛ' AND pvEndPokPUNomer.periodyear = '" + peryearprev + "' AND pvEndPokPUNomer.periodmonth = '" + permonthprev + "' AND pvEndPokPUNomer.lspropertieid = '7' AND pvEndPokPUNomer.otdelenieid = '" + otdelenieid + "') pvEndPokPUNomer" +
                    " ON pvStartPok2.codeIESBK = pvEndPokPUNomer.codeIESBK" +
                    //----------------------


                    " WHERE pvStartPok2.lstypeid = 'ФЛ' AND pvStartPok2.periodyear = '" + peryear + "' AND pvStartPok2.periodmonth = '" + permonth + "' AND pvStartPok2.lspropertieid = '22' AND pvStartPok2.otdelenieid = '" + otdelenieid + "'" +
                    " AND pvStartPok2.propvalue <> pvEndPok2.propvalue";

                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                //-----------------------------------------------------------               
                
                for (int j = 0; j < tableTOTALls.Rows.Count; j++)
                {
                    worksheet[strow + j + 1, stcol + 0].SetValue(tableTOTALls.Rows[j]["codels"].ToString());

                    worksheet[strow + j + 1, stcol + 1].SetValue(tableTOTALls.Rows[j]["pvStartPokFIOvalue"]);

                    worksheet[strow + j + 1, stcol + 2].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr1value"]);
                    worksheet[strow + j + 1, stcol + 3].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr2value"]);
                    worksheet[strow + j + 1, stcol + 4].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr3value"]);
                    worksheet[strow + j + 1, stcol + 5].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr4value"]);
                    worksheet[strow + j + 1, stcol + 6].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr5value"]);

                    // информация о предыдущем периоде
                    string date_prev_str = tableTOTALls.Rows[j]["prevPOKdatePrevPer"].ToString();                    
                    worksheet[strow + j + 1, stcol + 7].SetValue(date_prev_str);
                    worksheet[strow + j + 1, stcol + 7].Font.Color = infocolor_prev;

                    string value_prev_str = tableTOTALls.Rows[j]["prevPOKvaluePrevPer"].ToString();
                    worksheet[strow + j + 1, stcol + 8].SetValue(value_prev_str);
                    worksheet[strow + j + 1, stcol + 8].Font.Color = infocolor_prev;

                    string nomerPU_prev_str = tableTOTALls.Rows[j]["prevPokPUNomervalue"].ToString();
                    worksheet[strow + j + 1, stcol + 9].SetValue(nomerPU_prev_str);
                    worksheet[strow + j + 1, stcol + 9].Font.Color = infocolor_prev;

                    worksheet[strow + j + 1, stcol + 10].SetValue(tableTOTALls.Rows[j]["prevPOKtypePrevPer"]);
                    worksheet[strow + j + 1, stcol + 10].Font.Color = infocolor_prev;
                    //-------------------------------

                    // информация о текущем периоде
                    string date_tek_str = tableTOTALls.Rows[j]["startPOKdateTekPer"].ToString();
                    worksheet[strow + j + 1, stcol + 11].SetValue(date_tek_str);

                    string value_tek_str = tableTOTALls.Rows[j]["startPOKvalueTekPer"].ToString();
                    worksheet[strow + j + 1, stcol + 12].SetValue(value_tek_str);

                    string nomerPU_tek_str = tableTOTALls.Rows[j]["startPokPUNomervalue"].ToString();
                    worksheet[strow + j + 1, stcol + 13].SetValue(nomerPU_tek_str);

                    worksheet[strow + j + 1, stcol + 14].SetValue(tableTOTALls.Rows[j]["startPOKtypeTekPer"]);
                    //-------------------------------

                    worksheet[strow + j + 1, stcol + 8].NumberFormat = "#####";
                    worksheet[strow + j + 1, stcol + 12].NumberFormat = "#####";

                    // формируем и выводим статусы                 
                    bool flag_value_error = false; // нарушение нарастающего значения
                    bool flag_date_error = false; // нарушение нарастающей даты снятия показания

                    // если номер ПУ одинаковый
                    if (nomerPU_tek_str == nomerPU_prev_str)
                    {
                        if (date_tek_str != date_prev_str) flag_date_error = true;
                        if (value_tek_str != value_prev_str) flag_value_error = true;
                    }

                    if (flag_value_error) worksheet[strow + j + 1, stcol + 15].SetValue("да");
                    worksheet[strow + j + 1, stcol + 15].Font.Color = Color.Red;

                    if (flag_date_error) worksheet[strow + j + 1, stcol + 16].SetValue("да");
                    worksheet[strow + j + 1, stcol + 16].Font.Color = Color.Blue;
                    //----------------------------
                }

                // форматируем строку-заголовок
                worksheet.Rows[strow].Font.Bold = true;
                worksheet.Rows[strow].Alignment.WrapText = true;
                worksheet.Rows[strow].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[strow].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[strow].AutoFit();

                worksheet.Columns.AutoFit(0, MAXCOLinWRKSH);
                worksheet.Columns.Group(2, 6, true);
                worksheet.FreezeRows(strow);                
            } // for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)

            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
            splashScreenManager1.CloseWaitForm();

            form1.spreadsheetControl1.EndUpdate();

            form1.Show();
        } // ИЭСБК отчеты - "Нарушение нарастающего итога"

        public static string MyFUNC_GetPropValueFromIESBKOLAP2(int IESBKlsid, string periodyear, string periodmonth, int lspropidglobal,
            SqlConnection connection/*, out string pvstr, out double? pvdouble, out DateTime? pvdate*/)
        {
            string result = null;

            // получаем основной тип свойства
            /*string queryString =
                    "SELECT valuetypeid " +
                    "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                    "WHERE lspropidglobal = " + lspropidglobal.ToString();
            DataTable tablelsPropGLOBAL = new DataTable();
            MyFUNC_SelectDataFromSQLwoutConnection(tablelsPropGLOBAL, connection, queryString);            */

            DataRow[] lsproprows = tablelsPropGLOBAL.Select("lspropidglobal = " + lspropidglobal.ToString());
            int valuetypeid = Convert.ToInt32(lsproprows[0]["valuetypeid"].ToString());
            //int valuetypeid = Convert.ToInt32(tablelsPropGLOBAL.Rows[0]["valuetypeid"].ToString());

            //tablelsPropGLOBAL.Dispose();
            //-------------------------------

            DateTime period = Convert.ToDateTime("01." + periodmonth + "." + periodyear);

            // определяем таблицу назначения
            string propvaluetabledest = null;
            if (valuetypeid == 1)
            {
                propvaluetabledest = "tblIESBKlspropvaluestr"; // текстовый
                //result = "text";
            }
            else if (valuetypeid == 2)
            {
                propvaluetabledest = "tblIESBKlspropvaluenum"; // числовой
                //result = "123";
            }
            else if (valuetypeid == 3)
            {
                propvaluetabledest = "tblIESBKlspropvaluedate"; // дата
                //result = "01.12.1983";
            }
            //----------------------------

            string queryStringPropValue = "SELECT propvalue " +
                                          "FROM iesbk2.dbo." + propvaluetabledest +
                                          " WHERE lspropidglobal = " + lspropidglobal.ToString() + " AND period = '" + period.ToShortDateString() + "' AND IESBKlsid = " + IESBKlsid.ToString();
            //" WHERE IESBKlsid = " + IESBKlsid.ToString() + " AND lspropidglobal = " + lspropidglobal.ToString() + " AND period = '" + period.ToShortDateString() +"'";
            DataTable tablePropValue = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(connection, tablePropValue, queryStringPropValue);

            //pvstr = null; pvdate = null; pvdouble = null; // первоначальное обнуление возвращаемых значений

            if (tablePropValue.Rows.Count != 0)
            {
                result = tablePropValue.Rows[0]["propvalue"].ToString(); // формируем текстовое представление возвращаемого значения

                /*if (valuetypeid == 1) // текстовый
                {
                    pvstr = result; 
                }
                else if (valuetypeid == 2) // числовой
                {
                    pvdouble = Convert.ToDouble(tablePropValue.Rows[0]["propvalue"]); 
                }
                else if (valuetypeid == 3) // дата
                {
                    pvdate = Convert.ToDateTime(tablePropValue.Rows[0]["propvalue"]);
                } */
            }
            else if (valuetypeid != 1) // если значение свойства не найдено в основной не текстовой таблице, то смотрим в текстовой
            {
                queryStringPropValue = "SELECT propvalue " +
                                       "FROM iesbk2.dbo.tblIESBKlspropvaluestr" +
                                       " WHERE lspropidglobal = " + lspropidglobal.ToString() + " AND period = '" + period.ToShortDateString() + "' AND IESBKlsid = " + IESBKlsid.ToString();
                DataTable tablePropValue2 = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(connection, tablePropValue2, queryStringPropValue);

                if (tablePropValue2.Rows.Count != 0) result = tablePropValue2.Rows[0]["propvalue"].ToString();

                tablePropValue2.Dispose();
            }

            tablePropValue.Dispose();

            return result;
        }

        /*// пока БЕСПОЛЕЗНАЯ функция
        // функция возврата текстовых значений года и месяца в смещении от заданных
        // параметр offset считается "в месяцах" (+/-)
        public static DateTime MyFUNC_GetPeriodValuesOffset(string year_tek, string month_tek, int offset)
        {
            DateTime dt_tek = Convert.ToDateTime("01." + month_tek + "." + year_tek);
            DateTime dt_new = dt_tek.AddMonths(offset);
                        
            return dt_new;
        }*/

        // тест "месяца"
        private void barButtonItem39_ItemClick(object sender, ItemClickEventArgs e)
        {
            //textBox1.Text = DateTime.Now.Month.ToString();

            /*SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            textBox1.Text = MyFUNC_GetPropValueFromIESBKOLAP("КООО0016410", "2016", "06", 30, SQLconnection);

            SQLconnection.Close();

            DateTime dt = DateTime.Now;            
            dt = dt.AddMonths(-1);
            textBox1.Text = dt.ToString();*/

            DateTime dt = Convert.ToDateTime("10.08.2015");
            DateTime dt2 = Convert.ToDateTime("01.04.2016");

            textBox1.Text = dt.CompareTo(dt2).ToString();
        }

        // "шахматка" по среднемесячному
        private void barButtonItem40_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Проверка среднемесячного начисления";
            IWorkbook workbook = form1.spreadsheetControl1.Document;

            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            // выборка ЛС за текущий период 2016 по текущему отделению и по тем л/с, где поле Среднемесячное заполнено ---------------------------------------            
            DateTime dt_IESBK_MIN = Convert.ToDateTime("01.01.2015"); // левая граница имеющихся данных в OLAP-кубе

            //int MAX_PERIOD_MONTH = 12;
            int MAX_PERIOD_MONTH = 10;

            string yeartek = "2018";
            string monthtek = "10";
            string queryStringsrmes =
                                 "SELECT otdelenieid,codeIESBK,propvalue " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                 "WHERE" +
                                 //" otdelenieid = 'ВО' AND" + // ТЕСТ !!!!
                                 //" codeIESBK = '35000000041' AND" +
                                 " periodyear = '" + yeartek + "' AND periodmonth = '" + monthtek + "' AND lspropertieid='30'" + " AND propvalue IS NOT NULL";
            DataTable tablePOsrmestekPERIOD = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tablePOsrmestekPERIOD, queryStringsrmes);
            //------------------------------------------------------------------------------------

            Worksheet worksheet = workbook.Worksheets[0];
            //worksheet.Name = periodyeartek + ", " + periodmonthtek;

            // задаем смещение от левого угла листа книги            
            int strow = 0;
            int stcol = 0;
                        
            worksheet[strow + 0, stcol + 0].SetValue("Отделение ИЭСБК");
            worksheet[strow + 0, stcol + 1].SetValue("Код л/с ИЭСБК");

            worksheet[strow + 0, stcol + 2].SetValue("ФИО");
            worksheet[strow + 0, stcol + 3].SetValue("Населенный пункт");
            worksheet[strow + 0, stcol + 4].SetValue("Улица");
            worksheet[strow + 0, stcol + 5].SetValue("Дом");
            worksheet[strow + 0, stcol + 6].SetValue("Номер квартиры");

            // "пробегаем" по всем лицевым счетам выборки текущего периода            
            int rd = 1;
            int columns_in_period = 10+1;

            for (int i = 0; i < tablePOsrmestekPERIOD.Rows.Count; i++)
            //for (int i = 0; i < 500; i++)
            {
                string otdelenieid = tablePOsrmestekPERIOD.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.FindByotdelenieid(otdelenieid)["captionotd"].ToString();
                string codels = tablePOsrmestekPERIOD.Rows[i]["codeIESBK"].ToString();

                worksheet[strow + rd + 0, stcol + 0].SetValue(captionotd); // отделение ИЭСБК
                worksheet[strow + rd + 0, stcol + 1].SetValue(codels); // код ИЭСБК л/с
                
                // "бежим" по периодам с 01 по 09 2016 по выборке текущего периода -------------------                
                for (int period_i = 1; period_i <= MAX_PERIOD_MONTH; period_i++) //!!!
                {
                    //string periodyeartek = "2016";
                    string periodyeartek = "2018";
                    string periodmonthtek = (period_i < 10) ? "0" + period_i.ToString() : period_i.ToString();
                    
                    // выводим справочную информации на основе данных начального периода интервала
                    if (period_i == MAX_PERIOD_MONTH)
                    {
                        // выбираем все свойства л/с периода из базы
                        string queryStringlsprop = "SELECT lspropertieid, propvalue " +
                                         "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                         "WHERE codeIESBK='" + codels + "' AND periodyear = '" + periodyeartek + "' AND periodmonth = '" + periodmonthtek + "'";
                        DataTable tableLSprop = new DataTable();
                        MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

                        // ФИО
                        DataRow[] lsproprows = tableLSprop.Select("lspropertieid = '5'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 2].SetValue(lsproprows[0]["propvalue"].ToString());

                        // населенный пункт
                        lsproprows = tableLSprop.Select("lspropertieid = '12'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 3].SetValue(lsproprows[0]["propvalue"].ToString());

                        // улица
                        lsproprows = tableLSprop.Select("lspropertieid = '13'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 4].SetValue(lsproprows[0]["propvalue"].ToString());

                        // дом
                        lsproprows = tableLSprop.Select("lspropertieid = '14'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 5].SetValue(lsproprows[0]["propvalue"].ToString());

                        // номер квартиры
                        lsproprows = tableLSprop.Select("lspropertieid = '15'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 6].SetValue(lsproprows[0]["propvalue"].ToString());

                        tableLSprop.Dispose();
                    }

                    // выводим заголовки столбцов инфо о периоде
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(periodmonthtek + " Дата ПредПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(periodmonthtek + " ПредПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(periodmonthtek + " Вид ПредПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(periodmonthtek + " Дата ПослПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(periodmonthtek + " ПослПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(periodmonthtek + " Вид ПослПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(periodmonthtek + " СрМес РАСЧ");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 14].SetValue(periodmonthtek + " СрМес РАСЧ Дельта дней");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 14].Font.Color = Color.Green;

                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 15].SetValue(periodmonthtek + " ПолОтп ИЭСБК");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 15].Font.Color = Color.Blue;

                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(periodmonthtek + " СрМес ИЭСБК");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Blue;

                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 17].SetValue(periodmonthtek + " Недоп ПО");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 17].Font.Color = Color.Red;

                    // СрМес ПО ИЭСБК текущего периода
                    string po_iesbk_str = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, periodyeartek, periodmonthtek, 27, SQLconnection);
                    string srmes_iesbk_str = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, periodyeartek, periodmonthtek, 30, SQLconnection);
                    double? srmes_iesbk = null;
                    double? po_iesbk = null;

                    if (!String.IsNullOrWhiteSpace(po_iesbk_str) && !po_iesbk_str.Contains(";"))
                    {
                        po_iesbk = Convert.ToDouble(po_iesbk_str);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 15].Font.Color = Color.Blue;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 15].SetValue(po_iesbk); // Полезный отпуск ИЭСБК
                    }

                    if (!String.IsNullOrWhiteSpace(srmes_iesbk_str) && !srmes_iesbk_str.Contains(";"))
                    {
                        srmes_iesbk = Convert.ToDouble(srmes_iesbk_str);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Blue;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(srmes_iesbk); // СрМес ИЭСБК
                    }

                    // ищем ближайшее "правое" показание                
                    string value_right = null;
                    string dtvalue_right = null;

                    DateTime dt_right = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);
                    string year_right = null;
                    string month_right = null;

                    dt_right = dt_right.AddMonths(+1); // учитываем текущий месяц, т.е. +1-1 = 0

                    while (String.IsNullOrWhiteSpace(value_right) && dt_right >= dt_IESBK_MIN)
                    {
                        dt_right = dt_right.AddMonths(-1);
                        year_right = dt_right.Year.ToString();
                        month_right = null;
                        if (dt_right.Month < 10) month_right = "0" + dt_right.Month.ToString();
                        else month_right = dt_right.Month.ToString();

                        value_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 25, SQLconnection); // свойство "Текущее показание ПУ"                                        
                    };
                    //----------------------------------

                    // ищем "левое" показание, при условии, что нашли "правое" ------------------
                    string value_left = null;
                    string dtvalue_left = null;
                    string year_left = null;
                    string month_left = null;

                    if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    {
                        // получаем дату "правого" показания
                        dtvalue_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                        // выводим информацию о "правом" показании
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(dtvalue_right); // дата

                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(value_right); // показание

                        string rightpok_type = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 26, SQLconnection); 
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(rightpok_type); // вид

                        DateTime dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-5); // БЫЛО -6, отматываем 6 мес. (-5-1 = -6) = 180 дней от "правого" показания

                        while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN)
                        {
                            dt_left = dt_left.AddMonths(-1);
                            year_left = dt_left.Year.ToString();
                            month_left = null;
                            if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                            else month_left = dt_left.Month.ToString();

                            value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                        };

                        // если нет данных за период не менее 6 мес., то ищем за в периоде [6 мес.;3 мес.]
                        if (String.IsNullOrWhiteSpace(value_left))
                        {
                            dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-7); // отматываем 7 мес., т.к. в теле цикла сразу +1, т.е. -7+1 = -6

                            DateTime dt_IESBK_left_MAX = Convert.ToDateTime(dtvalue_right).AddMonths(-3);

                            while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN && dt_left < dt_IESBK_left_MAX)
                            {
                                dt_left = dt_left.AddMonths(+1);
                                year_left = dt_left.Year.ToString();
                                month_left = null;
                                if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                                else month_left = dt_left.Month.ToString();

                                value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                            };
                        } // if (String.IsNullOrWhiteSpace(value_left)) // если нет данных за период не менее 6 мес.

                        // получаем дату "левого" показания и выводим информацию о нем
                        if (!String.IsNullOrWhiteSpace(value_left))
                        {
                            dtvalue_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                            // выводим информацию о "правом" показании
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(dtvalue_left); // дата

                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(value_left); // показание

                            string leftpok_type = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 26, SQLconnection);
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(leftpok_type); // вид
                        }

                        // если даты "левого" и "правого" показаний не пустые, то формируем расчет среднемесячного
                        if (!String.IsNullOrWhiteSpace(dtvalue_left) && !dtvalue_left.Contains(";") && !String.IsNullOrWhiteSpace(dtvalue_right))
                        {
                            double pokleft = Convert.ToDouble(value_left);
                            double pokright = Convert.ToDouble(value_right);

                            // если не нарушен нарастающий итог
                            if (pokleft <= pokright)
                            {
                                System.TimeSpan deltaday = Convert.ToDateTime(dtvalue_right) - Convert.ToDateTime(dtvalue_left);
                                double deltapok = pokright - pokleft;

                                double srednesut_calc = deltapok / deltaday.Days;
                                double srmes_calc = Math.Round(srednesut_calc * DateTime.DaysInMonth(Convert.ToInt32(periodyeartek), Convert.ToInt32(periodmonthtek)));

                                // формируем отчет -----------------------------------------------
                                                                
                                if (!String.IsNullOrWhiteSpace(srmes_iesbk_str)) // еслм СрМес ПО ИЭСБК присутствует
                                {
                                    DateTime dt_period = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);

                                    if (srmes_iesbk >= 0 && Convert.ToDateTime(dtvalue_right).CompareTo(dt_period) < 0) // не выводим наши расчеты, если СрМес ИЭСБК < 0 и правое показание принадлежит текущему периоду анализа
                                    {
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(srmes_calc); // СрМес РАСЧ                                

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 14].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 14].SetValue(deltaday.Days); // СрМес РАСЧ Дельта дней

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 17].Font.Color = Color.Red;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 17].SetValue(srmes_calc - srmes_iesbk); // Недополученный ПО                                
                                    }
                                }
                                                                
                            } // if (pokleft <= pokright)
                              //-----------------------------------------------------------------
                        } // if (!String.IsNullOrWhiteSpace(dtvalue_left) && !String.IsNullOrWhiteSpace(dtvalue_right))


                    } // if (!String.IsNullOrWhiteSpace(value_right))
                      //----------------------------------
                      
                    splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " из " + tablePOsrmestekPERIOD.Rows.Count.ToString() + ")" + "-период " + period_i.ToString());

                }  // for (int period_i = 1; period_i < 7; period_i++)
                
                rd++;

            } // for (int i = 0; i < tablePOsrmestekPERIOD.Rows.Count; i++)

            //workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];

            worksheet.Columns.AutoFit(0, 7+10 * MAX_PERIOD_MONTH); // ПЕРЕДЕЛАТЬ для универсальности
            form1.spreadsheetControl1.EndUpdate();

            SQLconnection.Close();
            tablePOsrmestekPERIOD.Dispose();
            splashScreenManager1.CloseWaitForm();

            form1.Show();
        } // "шахматка" по среднемесячному

        // отчет на "разрыв" показаний (модернизированный "Нарушение нарастающего итога"), когда последние показания предыдущего расчетного периода не равны начальным текущего
        private void barButtonItem41_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Нарушение нарастающего итога";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            
            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            //for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            for (int i = 0; i < 1; i++)
            {
                string otdelenieid = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();

                Worksheet worksheet = workbook.Worksheets[i];
                worksheet.Name = captionotd;
                workbook.Worksheets.Add();

                // задаем смещение табличной части отчета
                int strow = 5;
                int stcol = 0;

                string peryearprev = "2016";
                string permonthprev = "06";

                string peryeartek = "2016";
                string permonthtek = "07";

                worksheet[0, 0].SetValue("ФЛ без ОДПУ");
                worksheet[1, 0].SetValue(captionotd);
                worksheet[2, 0].SetValue("Текущий период:");
                worksheet[2, 1].SetValue(peryeartek + ", " + permonthtek);
                worksheet[3, 0].SetValue("Предыдущий период:");
                worksheet[3, 1].SetValue(peryearprev + ", " + permonthprev);

                worksheet[strow + 0, stcol + 0].SetValue("Код ИЭСБК");
                worksheet[strow + 0, stcol + 1].SetValue("Пред период посл пок дата");
                worksheet[strow + 0, stcol + 2].SetValue("Пред период посл пок значение");
                worksheet[strow + 0, stcol + 3].SetValue("Пред период посл пок источник");
                worksheet[strow + 0, stcol + 4].SetValue("Тек период нач пок дата");
                worksheet[strow + 0, stcol + 5].SetValue("Тек период нач пок значение");
                worksheet[strow + 0, stcol + 6].SetValue("Тек период нач пок источник");

                worksheet[strow + 0, stcol + 7].SetValue("Нарушение по дате");
                worksheet[strow + 0, stcol + 8].SetValue("Разница показаний (НЕ ДОЛЖНО БЫТЬ)");

                string queryString =
                                 "SELECT otdelenieid,codeIESBK,propvalue " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                 "WHERE " +
                                 "periodyear = '" + peryeartek + "' AND periodmonth = '" + permonthtek + "'" + " AND otdelenieid = '" + otdelenieid + "' AND lspropertieid = '3'" + " AND propvalue IS NOT NULL";
                
                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                //-----------------------------------------------------------               
                int row_rep = 0;
                for (int j = 0; j < tableTOTALls.Rows.Count; j++)                
                {
                    string codeIESBK = tableTOTALls.Rows[j]["codeIESBK"].ToString();
                    
                    // информация о предыдущем периоде
                    string tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryearprev, permonthprev, 24, SQLconnection); // дата последнего показания периода
                    DateTime? POKprev_date_last = null;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKprev_date_last = Convert.ToDateTime(tempstr);
                    
                    tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryearprev, permonthprev, 25, SQLconnection); // значение последнего показания периода
                    double POKprev_value_last = -1;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKprev_value_last = Convert.ToDouble(tempstr);
                    
                    string POKprev_type_last = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryearprev, permonthprev, 26, SQLconnection); // вид последнего показания периода                    
                    

                    // информация о текущем периоде
                    tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryeartek, permonthtek, 21, SQLconnection); // дата предыдущего показания периода
                    DateTime? POKtek_date_first = null;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKtek_date_first = Convert.ToDateTime(tempstr);
                    
                    tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryeartek, permonthtek, 22, SQLconnection); // значение предыдущего показания периода
                    double POKtek_value_first = -1;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKtek_value_first = Convert.ToDouble(tempstr);
                    
                    string POKtek_type_first = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryeartek, permonthtek, 23, SQLconnection); // вид предыдущего показания периода                    

                    // нарушение может еще быть и по дате!!!!! ИСПРАВИТЬ
                    // + замена ПУ (тоже учесть, на всякий случай (кроме ";")

                    if (POKprev_value_last != POKtek_value_first && POKprev_value_last != -1 && POKtek_value_first != -1)
                    {
                        worksheet[strow + row_rep + 1, stcol + 0].SetValue(codeIESBK);

                        worksheet[strow + row_rep + 1, stcol + 1].SetValue(POKprev_date_last);
                        worksheet[strow + row_rep + 1, stcol + 2].SetValue(POKprev_value_last);
                        worksheet[strow + row_rep + 1, stcol + 3].SetValue(POKprev_type_last);

                        worksheet[strow + row_rep + 1, stcol + 4].SetValue(POKtek_date_first);
                        worksheet[strow + row_rep + 1, stcol + 5].SetValue(POKtek_value_first);
                        worksheet[strow + row_rep + 1, stcol + 6].SetValue(POKtek_type_first);

                        worksheet[strow + row_rep + 1, stcol + 7].SetValue(POKtek_date_first < POKprev_date_last ? "да" : null);
                        double deltaPOK = POKtek_value_first - POKprev_value_last;
                        worksheet[strow + row_rep + 1, stcol + 8].SetValue(deltaPOK);

                        row_rep++;
                    }

                    splashScreenManager1.SetWaitFormDescription("Обработка (" + (i+1).ToString() + " - " + (j + 1).ToString() + " из " + tableTOTALls.Rows.Count.ToString() + ")");
                }

                worksheet.Columns.AutoFit(0, 10+2);
            } // for (int j = 0; j < tableTOTALls.Rows.Count; j++)

            SQLconnection.Close();

            form1.spreadsheetControl1.EndUpdate();

            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // перенос данных "Свойства"
        private void barButtonItem42_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();
            
            // читаем "старые" свойства
            string queryStringlsprop = "SELECT lspropertieid, templateid, numcolumninfile, captionlsprop " +
                                        "FROM [iesbk].[dbo].[tblIESBKlsprop]";
            DataTable tableLSprop = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

            for (int i = 0; i < tableLSprop.Rows.Count; i++)
            {
                int templateid = 1;
                int lspropidtmpl = Convert.ToInt32(tableLSprop.Rows[i]["lspropertieid"].ToString());
                int lspropidglobal = templateid * 1000 + lspropidtmpl;

                string numcolumninfilestr = tableLSprop.Rows[i]["numcolumninfile"].ToString();
                int? numcolumninfile = null;
                if (!String.IsNullOrWhiteSpace(numcolumninfilestr)) numcolumninfile = Convert.ToInt32(numcolumninfilestr);
                                
                string captionlsprop = tableLSprop.Rows[i]["captionlsprop"].ToString();                

                string queryStringlsprop2 = "INSERT INTO iesbk2.dbo.tblIESBKlsprop(lspropidglobal, templateid, lspropidtmpl, numcolumninfile, captionlsprop, comment, valuetypeid) " +
                    "VALUES (" +
                    lspropidglobal.ToString() + "," +
                    templateid.ToString() + "," +
                    lspropidtmpl.ToString() + "," +                    
                    "NULL" + "," +
                    "'" + captionlsprop + "'" + "," +
                    "NULL" + "," + 
                    "-100" + ")";                    
                MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);

                /*queryStringlsprop2 = "UPDATE[iesbk2].[dbo].[tblIESBKlsprop]";
                MyFUNC_RunSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);*/

                splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " из " + tableLSprop.Rows.Count.ToString() + ")");
            }

            tableLSprop.Dispose();

            SQLconnection.Close();
            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Свойства"

        // перенос данных "Лицевые счета"
        private void barButtonItem43_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            // переносим по отделениям
            for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            {
                string otdelenieidstr = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();

                string queryString =
                    "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded " +
                    "FROM [iesbk].[dbo].[tblIESBKls] " +
                    "WHERE otdelenieid = '" + otdelenieidstr.ToString() + "'";
                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                //-----------------------------------------------------------               

                string queryString2 =
                    "SELECT otdelenieid, captionotd " +
                    "FROM [iesbk2].[dbo].[tblIESBKotdelenie] " +
                    "WHERE captionotd = '" + captionotd + "'";
                DataTable table2 = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(table2, dbconnectionStringIESBK2, queryString2);
                int otdelenieid = Convert.ToInt32(table2.Rows[0]["otdelenieid"].ToString());
                table2.Dispose();

                int lstypeid = 1; // фл без одпу
                int j = 0;
                for (j = 0; j < tableTOTALls.Rows.Count; j++)
                {
                    string codeIESBK = tableTOTALls.Rows[j]["codeIESBK"].ToString();
                    
                    DateTime dateloaded = Convert.ToDateTime(tableTOTALls.Rows[j]["dateloaded"].ToString());
                    int IESBKlsid = lstypeid * 100000000 + otdelenieid * 1000000 + j + 1;
                    //------------------------

                    string queryStringlsprop2 = "INSERT INTO iesbk2.dbo.tblIESBKls(codeIESBK, lstypeid, otdelenieid, dateloaded, IESBKlsid) " +
                        "VALUES (" +
                        "'" + codeIESBK + "'" + "," +
                        lstypeid.ToString() + "," +
                        otdelenieid.ToString() + "," +
                        "'" + dateloaded.ToString() + "'" + "," +
                        IESBKlsid.ToString() + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);

                    splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " - " + (j + 1).ToString() + " из " + tableTOTALls.Rows.Count.ToString() + ")");
                }

                // записываем последний lastlsnumber в таблицу [tblIESBKlastlsid]
                int lastlsid = lstypeid * 100000000 + otdelenieid * 1000000 + j;
                queryString2 = "UPDATE iesbk2.dbo.tblIESBKlastlsid " +
                                "SET lastlsid = " + lastlsid.ToString() +
                                "WHERE otdelenieid = " + otdelenieid.ToString();
                MC_SQLDataProvider.UpdateSQLQuery(dbconnectionStringIESBK2, queryString2);
                //-------------------------------------------------------------------

                tableTOTALls.Dispose();

            } // for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)

            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Лицевые счета"

        //-----------------------------------------------

        // перенос данных "Значения свойств"
        private void barButtonItem44_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            // параметры фильтра для "многозадачности" (запуск нескольких приложений)
            string peryear = textBox1.Text;
            string permonth = textBox3.Text;

            // формируем выборку лицевых счетов
            string queryString =
                    "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded " +
                    "FROM [iesbk].[dbo].[tblIESBKls]";                    
            DataTable tableTOTALls = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

            // "бежим" по выборке
            for (int i = 0; i < tableTOTALls.Rows.Count; i++)
            {
                string codeIESBK = tableTOTALls.Rows[i]["codeIESBK"].ToString();
                string otdelenieid_old = tableTOTALls.Rows[i]["otdelenieid"].ToString();

                // собираем все свойства "старого" лицевого счета через фильтр Код+Отделение+Год+Месяц (формируем фильтр-выборку)
                queryString =
                    "SELECT propvalue, periodyear, periodmonth, codeIESBK, lstypeid, lspropertieid, templateid, otdelenieid " +
                    "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                    "WHERE codeIESBK = '" + codeIESBK + "' AND otdelenieid = '" + otdelenieid_old + "'"
                    + " AND periodyear = '" + peryear + "'" + " AND periodmonth = '" + permonth + "'";
                DataTable tablelsprop = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsprop, dbconnectionStringIESBK, queryString);

                // получаем "новый" id отделения
                int otdelenieid_new = -1;
                if (otdelenieid_old == "АО") otdelenieid_new = 10;
                else if (otdelenieid_old == "ВО") otdelenieid_new = 11;
                else if (otdelenieid_old == "КО") otdelenieid_new = 12;
                else if (otdelenieid_old == "МЧО") otdelenieid_new = 13;
                else if (otdelenieid_old == "СлО") otdelenieid_new = 14;
                else if (otdelenieid_old == "СОЗ") otdelenieid_new = 15;
                else if (otdelenieid_old == "СОС") otdelenieid_new = 16;
                else if (otdelenieid_old == "ТОН") otdelenieid_new = 17;
                else if (otdelenieid_old == "ТОТ") otdelenieid_new = 18;
                else if (otdelenieid_old == "ТшО") otdelenieid_new = 19;
                else if (otdelenieid_old == "УКО") otdelenieid_new = 20;
                else if (otdelenieid_old == "УсО") otdelenieid_new = 21;
                else if (otdelenieid_old == "ЦО") otdelenieid_new = 22;
                else if (otdelenieid_old == "ЧО") otdelenieid_new = 23;
                //-------------------------------------------------

                // получаем id "нового" лицевого счета (учитывая отделение)
                queryString =
                    "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded, IESBKlsid "+
                    "FROM [iesbk2].[dbo].[tblIESBKls] " +
                    "WHERE codeIESBK = '" + codeIESBK + "' AND otdelenieid = '" + otdelenieid_new.ToString() + "'";
                DataTable tablelsnew = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                int IESBKlsid = Convert.ToInt32(tablelsnew.Rows[0]["IESBKlsid"].ToString());
                tablelsnew.Dispose();
                //-------------------------------------------------                

                // переносим свойства лицевого счета из фильтр-выборки
                for (int j = 0; j < tablelsprop.Rows.Count; j++)
                {                    
                    int templateid = 1; // 3.3.2
                    DateTime period = Convert.ToDateTime("01." + tablelsprop.Rows[j]["periodmonth"].ToString() + "." + tablelsprop.Rows[j]["periodyear"].ToString());

                    int lspropertieid = Convert.ToInt32(tablelsprop.Rows[j]["lspropertieid"].ToString());
                    string propvaluestr_src = tablelsprop.Rows[j]["propvalue"].ToString();

                    // получаем "новый" тип свойства
                    queryString =
                        "SELECT lspropidglobal, valuetypeid " +
                        "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                        "WHERE templateid = " + templateid.ToString() +
                        " AND lspropidtmpl = " + lspropertieid.ToString();
                    tablelsnew = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                    int valuetypeid = Convert.ToInt32(tablelsnew.Rows[0]["valuetypeid"].ToString());
                    int lspropidglobal = Convert.ToInt32(tablelsnew.Rows[0]["lspropidglobal"].ToString());
                    tablelsnew.Dispose();
                    //------------------------------

                    // если в значении свойства присутствует символ ";" (замена ПУ), то 
                    // загружаем в текстовый куб
                    if (propvaluestr_src.Contains(";")) valuetypeid = 1;

                    // очистка старых косяков
                    if (propvaluestr_src.Contains("REF!")) propvaluestr_src = null;
                    //-----------------------------------------------------------------

                    string propvaluetabledest = null;
                    if (valuetypeid == 1) propvaluetabledest = "tblIESBKlspropvaluestr"; // текстовый
                    else if (valuetypeid == 2) propvaluetabledest = "tblIESBKlspropvaluenum"; // числовой
                    else if (valuetypeid == 3) propvaluetabledest = "tblIESBKlspropvaluedate"; // дата
                    //----------------------                    

                    string propvaluestr_dest_for_insert = "NULL";
                    if (!String.IsNullOrWhiteSpace(propvaluestr_src))
                    {
                        if (valuetypeid == 1) propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                        else if (valuetypeid == 2)
                        {
                            //propvaluestr_dest_for_insert = Convert.ToDouble(propvaluestr_src).ToString().Replace(",","."); //? зачем конвертить

                            /*double doubleValue;                            
                            if (Double.TryParse(propvaluestr_src, out doubleValue)) propvaluestr_dest_for_insert = propvaluestr_src.Replace(",", ".");
                            else propvaluestr_dest_for_insert = "NULL";*/ // пока парсинг убрал - для увеличения производительности, при загрузке нужен обязательно!!!

                            propvaluestr_dest_for_insert = propvaluestr_src.Replace(",", ".");
                        }
                        else if (valuetypeid == 3)
                        {
                            //propvaluestr_dest_for_insert = "'" + Convert.ToDateTime(propvaluestr_src).ToShortDateString() + "'"; //? зачем конвертить

                            /*DateTime dateValue;
                            if (DateTime.TryParse(propvaluestr_src, out dateValue)) propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                            else propvaluestr_dest_for_insert = "NULL";*/

                            propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                        }
                    }

                    string queryStringlsprop2 = "INSERT INTO iesbk2.dbo." + propvaluetabledest + "(propvalue, lspropidglobal, period, IESBKlsid) " +
                        "VALUES (" +
                        propvaluestr_dest_for_insert + "," +
                        lspropidglobal.ToString() + "," +                        
                        "'" + period.ToShortDateString() + "'" + "," +
                        IESBKlsid.ToString() + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);

                    //splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " - " + (j + 1).ToString() + " из " + tablelsprop.Rows.Count.ToString() + ")");
                    splashScreenManager1.SetWaitFormDescription(peryear + "," + permonth + " (" + (i + 1).ToString() +  " из " + tableTOTALls.Rows.Count.ToString() + ")");
                }
                                
                tablelsprop.Dispose();

            } // for (int i = 0; i < tableTOTALls.Rows.Count; i++)

            tableTOTALls.Dispose();

            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Значения свойств"

        //----------------------------------------------

        // перенос данных "Значения свойств2"
        // - грузим полную выборку свойств - вывалилась с ошибкой переполнения
        // пробуем по периодам - год, месяц - долго (лучше, чем полная задница)
        // пробуем по отборам Код л/с / Код свойства - полная задница!!!
        // оставим по коду свойства
        private void barButtonItem45_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            //string peryear = "2015"; // ручной ввод года
            string peryear = textBox1.Text; // ручной ввод года

            int startlspropidtmpl = Convert.ToInt32(textBox2.Text);

            // считываем значение lastpropvalueid
            string queryString = "SELECT lastpropvalueid FROM [iesbk2].[dbo].[tblIESBKsystem]";
            DataTable table_system = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(table_system, dbconnectionStringIESBK, queryString);

            Int64 propvalueid = Convert.ToInt64(table_system.Rows[0]["lastpropvalueid"].ToString());

            table_system.Dispose();
            //-----------------------------------------------

            for (int lspropidtmpl = startlspropidtmpl; lspropidtmpl <= 54; lspropidtmpl++)
            {
                //--------------------------------------
                /*string permonth = periodmonth.ToString();
                if (periodmonth < 10) permonth = "0" + permonth;*/
                                    
                string lspropidtmplstr = lspropidtmpl.ToString();
                //--------------------------------------

                // берем полную выборку (по lspropertieid)
                queryString =
                        "SELECT propvalue, periodyear, periodmonth, codeIESBK, lstypeid, lspropertieid, templateid, otdelenieid " +
                        "FROM [iesbk].[dbo].[tblIESBKlspropvalue] "
                        //+ "WHERE periodyear = '" + peryear + "' AND periodmonth = '" + permonth + "'";
                        + "WHERE lspropertieid = '" + lspropidtmplstr + "'" + " AND periodyear = '" + peryear + "'";
                DataTable tablelsprop = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsprop, dbconnectionStringIESBK, queryString);

                // получаем тип свойства
                int templateid = 1; // 3.3.2
                                    //int lspropertieid = Convert.ToInt32(tablelsprop.Rows[j]["lspropertieid"].ToString());

                queryString =
                    "SELECT lspropidglobal, valuetypeid " +
                    "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                    "WHERE templateid = " + templateid.ToString() +
                    " AND lspropidtmpl = " + lspropidtmpl;
                DataTable tablelsnew = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                int valuetypeid = Convert.ToInt32(tablelsnew.Rows[0]["valuetypeid"].ToString());
                int lspropidglobal = Convert.ToInt32(tablelsnew.Rows[0]["lspropidglobal"].ToString());
                tablelsnew.Dispose();
                //------------------------------

                for (int j = 0; j < tablelsprop.Rows.Count; j++)
                {
                    string codeIESBK = tablelsprop.Rows[j]["codeIESBK"].ToString();
                    string otdelenieid_old = tablelsprop.Rows[j]["otdelenieid"].ToString();

                    int otdelenieid_new = -1;
                    if (otdelenieid_old == "АО") otdelenieid_new = 10;
                    else if (otdelenieid_old == "ВО") otdelenieid_new = 11;
                    else if (otdelenieid_old == "КО") otdelenieid_new = 12;
                    else if (otdelenieid_old == "МЧО") otdelenieid_new = 13;
                    else if (otdelenieid_old == "СлО") otdelenieid_new = 14;
                    else if (otdelenieid_old == "СОЗ") otdelenieid_new = 15;
                    else if (otdelenieid_old == "СОС") otdelenieid_new = 16;
                    else if (otdelenieid_old == "ТОН") otdelenieid_new = 17;
                    else if (otdelenieid_old == "ТОТ") otdelenieid_new = 18;
                    else if (otdelenieid_old == "ТшО") otdelenieid_new = 19;
                    else if (otdelenieid_old == "УКО") otdelenieid_new = 20;
                    else if (otdelenieid_old == "УсО") otdelenieid_new = 21;
                    else if (otdelenieid_old == "ЦО") otdelenieid_new = 22;
                    else if (otdelenieid_old == "ЧО") otdelenieid_new = 23;

                    // получаем id "нового" лицевого счета
                    queryString =
                        "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded, IESBKlsid " +
                        "FROM [iesbk2].[dbo].[tblIESBKls] " +
                        "WHERE codeIESBK = '" + codeIESBK + "' AND otdelenieid = '" + otdelenieid_new.ToString() + "'";
                    tablelsnew = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                    int IESBKlsid = Convert.ToInt32(tablelsnew.Rows[0]["IESBKlsid"].ToString());
                    tablelsnew.Dispose();
                    //-----------------------------------------------------------               

                    /*for (int j = 0; j < tablelsprop.Rows.Count; j++)
                    {*/
                    DateTime period = Convert.ToDateTime("01." + tablelsprop.Rows[j]["periodmonth"].ToString() + "." + tablelsprop.Rows[j]["periodyear"].ToString());
                    string propvaluestr_src = tablelsprop.Rows[j]["propvalue"].ToString();

                    /*// получаем тип свойства
                    int templateid = 1; // 3.3.2
                    int lspropertieid = Convert.ToInt32(tablelsprop.Rows[j]["lspropertieid"].ToString());

                    queryString =
                        "SELECT valuetypeid " +
                        "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                        "WHERE templateid = " + templateid.ToString() +
                        " AND lspropertieid = " + lspropertieid;
                    tablelsnew = new DataTable();
                    SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                    int valuetypeid = Convert.ToInt32(tablelsnew.Rows[0]["valuetypeid"].ToString());
                    tablelsnew.Dispose();
                    //------------------------------*/

                    // если в значении свойства присутствует символ ";" (замена ПУ), то 
                    // загружаем в текстовый куб
                    if (propvaluestr_src.Contains(";")) valuetypeid = 1;

                    // очистка старых косяков
                    if (propvaluestr_src.Contains("REF!")) propvaluestr_src = null;
                    //-----------------------------------------------------------------
                    
                    string propvaluestr_dest_for_insert = "NULL";
                    if (!String.IsNullOrWhiteSpace(propvaluestr_src))
                    {
                        if (valuetypeid == 1) propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                        else 
                        if (valuetypeid == 2) // число
                        {
                            propvaluestr_dest_for_insert = Convert.ToDouble(propvaluestr_src).ToString().Replace(",", ".");

                            /*double newDouble;
                            try
                            {
                                newDouble = Convert.ToDouble(propvaluestr_src);
                                propvaluestr_dest_for_insert = newDouble.ToString().Replace(",", ".");
                            }                            
                            catch (System.FormatException) // если нарушен формат, то считаем текстом
                            {
                                valuetypeid = 1;
                                propvaluestr_dest_for_insert = propvaluestr_src;
                            }*/
                        }
                        else 
                        if (valuetypeid == 3) // дата
                        {
                            propvaluestr_dest_for_insert = "'" + Convert.ToDateTime(propvaluestr_src).ToShortDateString() + "'";

                            /*DateTime newDate;
                            try
                            {
                                newDate = Convert.ToDateTime(propvaluestr_src);
                                propvaluestr_dest_for_insert = "'" + newDate.ToShortDateString() + "'";
                            }
                            catch (System.FormatException) // если нарушен формат, то считаем текстом
                            {
                                valuetypeid = 1;
                                propvaluestr_dest_for_insert = propvaluestr_src;
                            }*/
                        }
                    }
                    
                    // определяем таблицу назначения
                    string propvaluetabledest = null;
                    if (valuetypeid == 1)
                    {
                        propvaluetabledest = "tblIESBKlspropvaluestr"; // текстовый
                    }
                    else if (valuetypeid == 2)
                    {
                        propvaluetabledest = "tblIESBKlspropvaluenum"; // числовой
                    }
                    else if (valuetypeid == 3)
                    {
                        propvaluetabledest = "tblIESBKlspropvaluedate"; // дата время
                    }

                    // вставляем запись в куб "ссылок" на значения свойств
                    propvalueid++; // нумерация с 1
                    string queryStringlspropvalue = "INSERT INTO iesbk2.dbo.tblIESBKlspropvalue(propvalueid, lspropidglobal, period, IESBKlsid, valuetypeid) " +
                        "VALUES (" +
                        propvalueid.ToString() + "," +
                        lspropidglobal.ToString() + "," +
                        "'" + period.ToShortDateString() + "'" + "," +
                        IESBKlsid.ToString() + "," +
                        valuetypeid.ToString() + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlspropvalue);

                    // вставляем значения свойств по ссылке из куба
                    string queryStringlspropvalue2 = "INSERT INTO iesbk2.dbo." + propvaluetabledest + "(propvalueid, propvalue) " +
                        "VALUES (" +
                        propvalueid.ToString() + "," +
                        propvaluestr_dest_for_insert + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlspropvalue2);

                    splashScreenManager1.SetWaitFormDescription("Обработка (" + lspropidtmplstr + " - " + (j + 1).ToString() + " из " + tablelsprop.Rows.Count.ToString() + ")");
                    //} // for (int j = 0; j < tablelsprop.Rows.Count; j++)

                    //tablelsprop.Dispose();

                } // for (int i = 0; i < tableTOTALls.Rows.Count; i++)

                tablelsprop.Dispose();

            } // for (int periodmonth = 1; periodmonth <= 12; periodmonth++)                

            // записываем последний lastpropvalueid в таблицу [tblIESBKsystem]            
            string queryString2 = "UPDATE iesbk2.dbo.tblIESBKsystem " +
                            "SET lastlsid = " + propvalueid.ToString();
            MC_SQLDataProvider.UpdateSQLQuery(dbconnectionStringIESBK2, queryString2);

            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Значения свойств"

        // отчет-"шахматка" по наличию л/с и полезного отпуска - 2
        private void barButtonItem46_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK2);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Шахматка с января 2016 по сентябрь 2016";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // константы
            int MAX_PERIOD_MONTH = 9;
            int columns_in_period_auto = 20; // колонок в периоде для автоматического вывода
            int columns_in_period_manual = 4; // колонок в периоде для ручного вывода
            int columns_in_period = columns_in_period_auto + columns_in_period_manual; // общее кол-во колонок в периоде
            int FIRST_COLUMNS = 8;
            int END_COLUMNS = 1 + 3 + 8 + MAX_PERIOD_MONTH;

            //int MAXPROPMAS = 9;
            int MAXCOLinWRKSH = FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + END_COLUMNS;

            for (int col = 0; col < MAXCOLinWRKSH; col++)
            {
                worksheet.Columns[col].Font.Name = "Arial";
                worksheet.Columns[col].Font.Size = 8;
            }

            DateTime dt_IESBK_MIN = Convert.ToDateTime("01.01.2015"); // левая граница имеющихся данных в OLAP-кубе

            // загружаем параметры свойств в глобальную таблицу (тестируем производительность)
            string queryString22 =
                    "SELECT lspropidglobal, valuetypeid " +
                    "FROM [iesbk2].[dbo].[tblIESBKlsprop]";
            tablelsPropGLOBAL = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tablelsPropGLOBAL, queryString22);            
            //-------------------------------

            /*// загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);*/
            string queryString = "SELECT otdelenieid, captionotd " +
                                 "FROM [iesbk2].[dbo].[tblIESBKotdelenie]";                                
            DataTable tableIESBKotdelenie = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableIESBKotdelenie, queryString);

            //-----------------------------------------

            // продумать выборку!!! пробуем JOIN

            // св-во 36 - "Расход ОДН по нормативу" (убрал)
            /*string queryString = "SELECT DISTINCT codeIESBK,otdelenieid " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +                                
                                "WHERE periodyear = '2016' AND (periodmonth = '01' OR periodmonth = '07') AND lspropertieid='36' AND propvalue IS NULL";*/

            /*queryString = "SELECT DISTINCT IESBKlsid " +
                          "FROM [iesbk2].[dbo].[tblIESBKlspropvaluestr] " +
                          "WHERE lspropidglobal = 1003 AND DATEPART(year, period) = 2016";*/

            queryString = "SELECT tblIESBKPVstr.IESBKlsid, tblIESBKls.codeIESBK, tblIESBKotd.captionotd" +
                          " FROM" + 
                          " (SELECT DISTINCT IESBKlsid, lspropidglobal" +
                          " FROM iesbk2.dbo.tblIESBKlspropvaluestr" +
                          " WHERE lspropidglobal = 1003 AND DATEPART(year, period) = 2016) tblIESBKPVstr" + 
                          " LEFT JOIN iesbk2.dbo.tblIESBKls tblIESBKls ON tblIESBKPVstr.IESBKlsid = tblIESBKls.IESBKlsid" +
                          " LEFT JOIN iesbk2.dbo.tblIESBKotdelenie tblIESBKotd ON tblIESBKls.otdelenieid = tblIESBKotd.otdelenieid";

            //+ " AND (otdelenieid = 'СОС' OR otdelenieid = 'СОЗ')";// AND lspropertieid='36' AND propvalue IS NULL";
            DataTable tableTOTALls10 = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls10, dbconnectionStringIESBK2, queryString);
            //-----------------------------

            // выводим заголовки столбцов -----------------------------------------------------

            // статические
            worksheet[0, 0].SetValue("№ п/п");
            worksheet[0, 1].SetValue("Отделение ИЭСБК");
            worksheet[0, 2].SetValue("Код л/с ИЭСБК");
            worksheet[0, 3].SetValue("ФИО");
            worksheet[0, 4].SetValue("Населенный пункт");
            worksheet[0, 5].SetValue("Улица");
            worksheet[0, 6].SetValue("Дом");
            worksheet[0, 7].SetValue("Номер квартиры");
            //worksheet[0, 8].SetValue("Состояние ЛС (на 2016 09)");

            // периодические
            // id "периодических" свойств
            //int[] propidmas = new int[] { 6, 7, 50, 24, 25, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };
            int[] propidmas = new int[] { 1051, 1024, 1025, 1006, 1007, 1050, 1026, 1027, 1028, 1055, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034, 1035, 1036 };
            int idprop_PO_in_propidmas = 7; // индекс идентификатора поля ПО от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_lastPOK_in_propidmas = 2; // индекс идентификатора поля ПослПоказаниеПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_nomerPU_in_propidmas = 4; // индекс идентификатора поля ЗаводскойНомерПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int[] propidmas_doublevalue = new int[] { 1027, 1028, 1055, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034, 1035, 1036 }; // id числовых полей

            string queryStringlsprop = "SELECT lspropidglobal, captionlsprop " +
                                        "FROM [iesbk2].[dbo].[tblIESBKlsprop]";
            DataTable tableLSprop = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

            Color Color_IS_PO_NORMATIV_NOT_PU = Color.Orange; // цвет "норматив - безприборник"
            Color Color_IS_PO_NORMATIV_YES_PU = Color.Red; // цвет "норматив - приборник"
            Color Color_IS_PO_SREDNEMES_YES_PU = Color.Blue; // цвет "среднемесячное - приборник"
            Color Color_IS_PO_RASHOD_YES_PU = Color.Green; // цвет "расход по прибору"

            for (int period_i = 1; period_i < MAX_PERIOD_MONTH + 1; period_i++)
            {
                string periodyear = "2016";
                string periodmonth = (period_i < 10) ? "0" + period_i.ToString() : period_i.ToString();

                for (int k = 0; k < columns_in_period_auto; k++)
                {
                    DataRow[] lsproprows = tableLSprop.Select("lspropidglobal = " + propidmas[k].ToString());
                    string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["captionlsprop"].ToString() : null;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(periodyear + " " + periodmonth + ", " + propvaluestr);

                    Color cellColor = Color.Black;
                    // "красим" названия столбцов
                    if (propidmas[k] == 1028) cellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (propidmas[k] == 1029) cellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (propidmas[k] == 1030 || propidmas[k] == 53) cellColor = Color_IS_PO_SREDNEMES_YES_PU;
                    else if (propidmas[k] == 1031 || propidmas[k] == 54) cellColor = Color_IS_PO_NORMATIV_YES_PU;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].Font.Color = cellColor;
                }

                tableLSprop.Dispose();

                //------------------------------------------------------------------------------

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(periodyear + " " + periodmonth + ", " + "Расчетный полезный отпуск");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(periodyear + " " + periodmonth + ", " + "Расход по показаниям ПУ");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(periodyear + " " + periodmonth + ", " + "Начисленный полезный отпуск от ИЭСБК");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(periodyear + " " + periodmonth + ", " + "Отклонение (недополученный ПО)");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;
            }

            //------------------------------------------------------------------------------
            string periodmonthnext = (MAX_PERIOD_MONTH + 1 < 10) ? "0" + (MAX_PERIOD_MONTH + 1).ToString() : (MAX_PERIOD_MONTH + 1).ToString();
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue("2016" + " " + periodmonthnext + ", " + "среднемесячное (прогноз ПО)");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue("ИТОГО Расход по показаниям ПУ с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].Font.Color = Color.Green;

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue("Начальное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue("Начальное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue("Начальное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue("Начальное показание, номер ПУ");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue("Конечное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue("Конечное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue("Конечное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue("Конечное показание, номер ПУ");

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue("ИТОГО Полезный отпуск от ИЭСБК с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].Font.Color = Color.Blue;

            for (int period_i = 1; period_i <= MAX_PERIOD_MONTH; period_i++)
            {
                worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10 + period_i].
                    SetValue(worksheet[0, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString());
            }

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].SetValue("ИТОГО Отклонение (недополученный ПО) с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].Font.Color = Color.Red;
            //----------------------------------------------------------

            // главный цикл
            //for (int i = 0; i < tableTOTALls10.Rows.Count; i++)
            for (int i = 0; i < 400; i++)
            {
                //DataRow[] lsproprows;

                // получаем внутренний идентификатор л/с, отделение и код ИЭСБК
                /*queryString = "SELECT IESBKlsid, codeIESBK, otdelenieid " +
                              "FROM [iesbk2].[dbo].[tblIESBKls] " +
                              "WHERE IESBKlsid = " + tableTOTALls10.Rows[i]["IESBKlsid"].ToString();
                DataTable tableIESBKls = new DataTable();
                MyFUNC_SelectDataFromSQLwoutConnection(tableIESBKls, SQLconnection, queryString);
                
                int IESBKlsid = Convert.ToInt32(tableTOTALls10.Rows[i]["IESBKlsid"]);
                string codeIESBK = tableIESBKls.Rows[0]["codeIESBK"].ToString();
                //string codeIESBK = "ККОО00019257";
                string otdelenieid = tableIESBKls.Rows[0]["otdelenieid"].ToString();
                
                tableIESBKls.Dispose();

                string otdeleniecapt = tableIESBKotdelenie.Select("otdelenieid = " + otdelenieid)[0]["captionotd"].ToString();*/
                //string otdeleniecapt = "АО";
                                
                int IESBKlsid = Convert.ToInt32(tableTOTALls10.Rows[i]["IESBKlsid"]);
                string codeIESBK = tableTOTALls10.Rows[i]["codeIESBK"].ToString();                
                string otdeleniecapt = tableTOTALls10.Rows[i]["captionotd"].ToString();
                //--------------------------------

                worksheet[i + 1, 0].SetValue((i + 1).ToString());
                worksheet[i + 1, 1].SetValue(otdeleniecapt);
                worksheet[i + 1, 2].SetValue(codeIESBK);

                /*// состояние ЛС за июнь 2016  ---------------------------------------
                string periodyear = "2016";
                string periodmonth = "06";                    
                worksheet[i + 1, 3].SetValue(MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 51, SQLconnection));*/

                // основные параметры ЛС за начальный период (01 2016)  ---------------------------------------
                string periodyear = "2016";
                string periodmonth = "01";

                // сделано для увеличения производительности
                /*// "статические" свойства лицевого счета                
                queryStringlsprop = "SELECT codeIESBK, lspropertieid, propvalue " +
                                     "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                     "WHERE codeIESBK='" + codeIESBK + "' AND periodyear = '" + periodyear + "' AND periodmonth = '" + periodmonth + "'";
                tableLSprop = new DataTable();
                MyFUNC_SelectDataFromSQLwoutConnection(tableLSprop, SQLconnection, queryStringlsprop);*/
                                                
                // фио
                /*string lspropvalue = "";
                                
                lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 5, SQLconnection);
                worksheet[i + 1, 3].SetValue(lspropvalue);

                /*DataRow[] lsproprows = tableLSprop.Select("lspropertieid='5'");
                if (lsproprows.Length > 0) worksheet[i + 1, 3].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // населенный пункт                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 12, SQLconnection);
                worksheet[i + 1, 4].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='12'");
                if (lsproprows.Length > 0) worksheet[i + 1, 4].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // улица                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 13, SQLconnection);
                worksheet[i + 1, 5].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='13'");
                if (lsproprows.Length > 0) worksheet[i + 1, 5].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // дом                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 14, SQLconnection);
                worksheet[i + 1, 6].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='14'");
                if (lsproprows.Length > 0) worksheet[i + 1, 6].SetValue(lsproprows[0]["propvalue"].ToString());                */

                // номер квартиры                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 15, SQLconnection);
                worksheet[i + 1, 7].SetValue(lspropvalue);                
                /*lsproprows = tableLSprop.Select("lspropertieid='15'");
                if (lsproprows.Length > 0) worksheet[i + 1, 7].SetValue(lsproprows[0]["propvalue"].ToString());                */

                //tableLSprop.Dispose();
                //---------------------------------------

                // переменные для подсчета ПО по разнице показаний + сам ПО
                double POKstart = -1;
                double POKend = -1;

                DateTime POKstart_date = Convert.ToDateTime("01.01.1900");
                DateTime POKend_date = Convert.ToDateTime("01.01.1900");
                string POKstart_kind = null;
                string POKend_kind = null;
                int periodPOKstart = -1; // нумерация с 1 (январь)
                int periodPOKend = -1; // нумерация с 1 (январь)

                string nomerPUstart = null;
                string nomerPUend = null;
                //------------------------------------------------

                // "бежим" по периодическим свойствам лицевого счета
                int START_PERIOD_MONTH = 1;
                for (int period_i = START_PERIOD_MONTH; period_i < START_PERIOD_MONTH + MAX_PERIOD_MONTH; period_i++)
                {
                    periodyear = "2016";
                    periodmonth = (period_i < 10) ? "0" + period_i.ToString() : period_i.ToString();

                    DateTime periodAsDateTime = Convert.ToDateTime("01." + periodmonth + "." + periodyear);

                    /*queryStringlsprop = "SELECT lspropidglobal, propvalue " +
                                     "FROM [iesbk2].[dbo].[tblIESBKlspropvaluestr] " +
                                     "WHERE period ='" + periodAsDateTime + "' AND IESBKlsid = " + IESBKlsid.ToString();
                    tableLSprop = new DataTable();
                    MyFUNC_SelectDataFromSQLwoutConnection(tableLSprop, SQLconnection, queryStringlsprop);                    */
                    //-----------------------------------
                    // "статические" свойства лицевого счета                                
                                        
                    string lspropvalue = "";

                    // фио                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1005, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 3].SetValue(lspropvalue);

                    // населенный пункт                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1012, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 4].SetValue(lspropvalue);

                    // улица                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1013, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 5].SetValue(lspropvalue);

                    // дом                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1014, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 6].SetValue(lspropvalue);

                    // номер квартиры                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1015, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 7].SetValue(lspropvalue);
                    //-----------------------------------
                                        
                    /*// состояние ЛС (по последнему расчетному периоду)
                    if (period_i == MAX_PERIOD_MONTH)
                    {                        
                        worksheet[i + 1, 8].SetValue(MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1051, SQLconnection));
                    }*/

                    // флаги для раскраски ячейки полезного отпуска                    
                    bool IS_PO_NORMATIV_NOT_PU = false; // флаг "норматив - безприборник"
                    bool IS_PO_NORMATIV_YES_PU = false; // флаг "норматив - приборник"
                    bool IS_PO_SREDNEMES_YES_PU = false; // флаг "среднемесячное - приборник"
                    bool IS_PO_RASHOD_YES_PU = false; // флаг "расход по прибору"

                    // выводим "периодические" поля - сделал в цикле                    
                    for (int k = 0; k < columns_in_period_auto; k++)
                    {
                        /*string strTEST = "12345678";
                        string propvaluestr = strTEST;*/

                        /*lsproprows = tableLSprop.Select("lspropertieid = '" + propidmas[k].ToString() + "'");
                        string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                        string propvaluestr = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, propidmas[k], SQLconnection);

                        double? propvalue = null;

                        if (System.Array.IndexOf(propidmas_doublevalue, propidmas[k]) >= 0)
                        {
                            if (!String.IsNullOrWhiteSpace(propvaluestr)) propvalue = Convert.ToDouble(propvaluestr);
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvalue);
                        }
                        else
                        {
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvaluestr);
                        }

                        // формируем флаги для раскраски ячейки полезного отпуска
                        // помним о массиве периодических свойств
                        //int[] propidmas = new int[] { 1006, 1007, 1050, 1024, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };                        
                        if (propvalue != null && propvalue != 0)
                        {
                            if (propidmas[k] == 1028) IS_PO_NORMATIV_NOT_PU = true;
                            else if (propidmas[k] == 1029) IS_PO_RASHOD_YES_PU = true;
                            else if (propidmas[k] == 1030) IS_PO_SREDNEMES_YES_PU = true;
                            else if (propidmas[k] == 1031) IS_PO_NORMATIV_YES_PU = true;
                        };
                        //-------------------------------------------------------
                    } // выводим "периодические" поля - сделал в цикле

                    // раскрашиваем колонку ПолОтп в зависимости от "слагаемых"
                    // помним о массиве периодических свойств
                    //int[] propidmas = new int[] { 1006, 1007, 1050, 1024, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };

                    Color PolOtpCellColor = Color.Black;
                    if (IS_PO_RASHOD_YES_PU) PolOtpCellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (IS_PO_NORMATIV_NOT_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (IS_PO_NORMATIV_YES_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_YES_PU;
                    else if (IS_PO_SREDNEMES_YES_PU) PolOtpCellColor = Color_IS_PO_SREDNEMES_YES_PU;
                    worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + idprop_PO_in_propidmas].Font.Color = PolOtpCellColor; // "красим" ПолезныйОтпуск - propid = 27

                    //-----------------------------------------------------------------
                    // формируем "периодические" колонки "Расход по показаниям" и "Отклонение (недополученный ПО)", если не было замены ПУ

                    if (period_i > START_PERIOD_MONTH) // пропускаем первый месяц, т.к. в нем не найдем "предыдущих показаний"
                    {
                        string propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        double POKend_period = -1;
                        //double? POIESBK_period = null;
                        double POIESBK_period = 0;
                        double POIESBKend_period = 0; // ПО последнего периода

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";"))
                        {
                            POKend_period = Convert.ToDouble(propvaluestr);

                            double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                            POIESBK_period += POcellvalue;

                            POIESBKend_period = POcellvalue;
                        }

                        string nomerPUend_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        /*lsproprows = tableLSprop.Select("lspropertieid='22'"); // предыдущее показание ПУ
                        propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                        int period_i_step = 1;
                        double POKstart_period = -1;
                        string nomerPUstart_period = null;
                        propvaluestr = null;

                        do
                        {
                            propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();
                            nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();

                            // суммируем полезный отпуск, пропуская начальный интервал
                            if (String.IsNullOrWhiteSpace(propvaluestr))
                            {
                                double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.NumericValue;
                                POIESBK_period += POcellvalue;
                            }

                            period_i_step += 1;

                        } while (String.IsNullOrWhiteSpace(propvaluestr) && period_i_step < period_i);

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POKstart_period = Convert.ToDouble(propvaluestr);

                        //nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 2) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        // если имеются оба показания и не было замены ПУ, то считаем значения
                        if (POKend_period != -1 && POKstart_period != -1 && String.Equals(nomerPUstart_period, nomerPUend_period))
                        //if (POKend_period != -1 && POKstart_period != -1 && POKstart_period <= POKend_period)                         
                        {
                            /*propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                            double? POIESBK_period = null;
                            if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POIESBK_period = Convert.ToDouble(propvaluestr);*/

                            double POIESBKPU_period = POKend_period - POKstart_period;
                            double POIESBKDelta_period = POIESBKPU_period - POIESBK_period;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(POIESBKend_period + POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(POIESBKPU_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(POIESBK_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;
                        }

                    } // if (period_i > START_PERIOD_MONTH) 
                    //-----------------------------------------------------------------

                    // ищем стартовое и конечное показание для расчета ИТОГОВОГО ПО по показаниям (за все периоды)                  
                    /*lsproprows = tableLSprop.Select("lspropertieid='25'");
                    string pokstr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                    string pokstr = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1025, SQLconnection);

                    if (!String.IsNullOrWhiteSpace(pokstr) && !pokstr.Contains(";"))
                    {
                        if (POKstart == -1)
                        {
                            POKstart = Convert.ToDouble(pokstr);
                            periodPOKstart = period_i;

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1007");
                            nomerPUstart = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            nomerPUstart = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1007, SQLconnection);

                            /*lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKstart_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());*/
                            POKstart_date = Convert.ToDateTime(MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1024, SQLconnection));

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1026");
                            POKstart_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            POKstart_kind = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1026, SQLconnection);
                        }
                        else
                        {
                            // поменять местами нижнее условие и присовение значение конечному показанию!!!!!
                            POKend = Convert.ToDouble(pokstr);
                            periodPOKend = period_i;

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1007");
                            nomerPUend = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            nomerPUend = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1007, SQLconnection);

                            /*lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKend_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());*/
                            POKend_date = Convert.ToDateTime(MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1024, SQLconnection));

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1026");
                            POKend_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            POKend_kind = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1026, SQLconnection);

                            if (!String.Equals(nomerPUstart, nomerPUend)) // если номера ПУ не равны, то последнее делаем начальным
                            {
                                nomerPUstart = nomerPUend;
                                POKstart = POKend;
                                POKstart_date = POKend_date;
                                POKstart_kind = POKend_kind;

                                POKend = -1;

                                periodPOKstart = periodPOKend;
                                periodPOKend = -1;
                            };
                        }
                    }

                    //-----------------------------------------------------------------

                    tableLSprop.Dispose();

                } // for (int period_i = 1; period_i < 7; period_i++)

                if (POKstart != -1 && POKend != -1) // если имеются оба показания ПУ (для расчета ПО по показаниям ПУ)
                {
                    /*// суммируем полезный отпуск от ИЭСБК со следующего расчетного периода, которому предшествовало показание ПУ
                    lsproprows = tableLSprop.Select("lspropertieid='27'");
                    string postr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                    if (POKstart != -1 && POKend == -1 && !String.IsNullOrWhiteSpace(postr)) POIESBKTotal += Convert.ToDouble(postr);*/

                    // суммируем полезный отпуск от ИЭСБК по ранее заполненным колонкам со следующего расчетного периода, которому предшествовало показание ПУ
                    double POIESBKTotal = 0;
                    for (int period_i = periodPOKstart + 1; period_i <= periodPOKend; period_i++)
                    {
                        double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                        POIESBKTotal += POcellvalue;

                        // выводим расшифровку формирования ПО
                        worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10 + period_i].SetValue(POcellvalue);
                    }

                    double POIESBKPU = POKend - POKstart;
                    double POIESBKDelta = POIESBKPU - POIESBKTotal;

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue(POIESBKPU);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].Font.Color = Color.Green;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue(POKstart_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue(POKstart);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue(POKstart_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue(nomerPUstart);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue(POKend_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue(POKend);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue(POKend_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue(nomerPUend);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue(POIESBKTotal);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].Font.Color = Color.Blue;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].SetValue(POIESBKDelta);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].Font.Color = Color.Red;

                    // beg расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)

                    string periodyeartek = "2016";
                    string periodmonthtek = (MAX_PERIOD_MONTH < 10) ? "0" + MAX_PERIOD_MONTH.ToString() : MAX_PERIOD_MONTH.ToString();
                    //string codels = codeIESBK;

                    // ищем ближайшее "правое" показание                
                    string value_right = null;
                    string dtvalue_right = null;

                    DateTime dt_right = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);
                    string year_right = null;
                    string month_right = null;

                    dt_right = dt_right.AddMonths(+1); // учитываем текущий месяц, т.е. +1-1 = 0

                    while (String.IsNullOrWhiteSpace(value_right) && dt_right >= dt_IESBK_MIN)
                    {
                        dt_right = dt_right.AddMonths(-1);
                        year_right = dt_right.Year.ToString();
                        month_right = null;
                        if (dt_right.Month < 10) month_right = "0" + dt_right.Month.ToString();
                        else month_right = dt_right.Month.ToString();

                        value_right = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_right, month_right, 1025, SQLconnection); // свойство "Текущее показание ПУ"                                        
                    };
                    //----------------------------------

                    // ищем "левое" показание, при условии, что нашли "правое" ------------------
                    string value_left = null;
                    string dtvalue_left = null;
                    string year_left = null;
                    string month_left = null;

                    if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    {
                        // получаем дату "правого" показания
                        dtvalue_right = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_right, month_right, 1024, SQLconnection); // свойство "Дата последнего показания ПУ"

                        /*// выводим информацию о "правом" показании
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(dtvalue_right); // дата

                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(value_right); // показание

                        string rightpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_right, month_right, 26, SQLconnection);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(rightpok_type); // вид*/

                        DateTime dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-5); // отматываем -5-1 = -6 мес. (было 6) = 180 дней от "правого" показания

                        while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN)
                        {
                            dt_left = dt_left.AddMonths(-1);
                            year_left = dt_left.Year.ToString();
                            month_left = null;
                            if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                            else month_left = dt_left.Month.ToString();

                            value_left = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_left, month_left, 1025, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                        };

                        // если нет данных да период не менее 6 мес., то ищем за в периоде [6 мес.;3 мес.]
                        if (String.IsNullOrWhiteSpace(value_left))
                        {
                            dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-7); // отматываем 7 мес., т.к. в теле цикла сразу +1, т.е. -7+1 = -6

                            DateTime dt_IESBK_left_MAX = Convert.ToDateTime(dtvalue_right).AddMonths(-3);

                            while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN && dt_left < dt_IESBK_left_MAX)
                            {
                                dt_left = dt_left.AddMonths(+1);
                                year_left = dt_left.Year.ToString();
                                month_left = null;
                                if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                                else month_left = dt_left.Month.ToString();

                                value_left = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_left, month_left, 1025, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                            };
                        } // if (String.IsNullOrWhiteSpace(value_left)) // если нет данных за период не менее 6 мес.

                        // получаем дату "левого" показания и выводим информацию о нем
                        if (!String.IsNullOrWhiteSpace(value_left))
                        {
                            dtvalue_left = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_left, month_left, 1024, SQLconnection); // свойство "Дата последнего показания ПУ"

                            /*// выводим информацию о "правом" показании
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(dtvalue_left); // дата

                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(value_left); // показание

                            string leftpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_left, month_left, 26, SQLconnection);
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(leftpok_type); // вид*/
                        }

                        // если даты "левого" и "правого" показаний не пустые, то формируем расчет среднемесячного
                        if (!String.IsNullOrWhiteSpace(dtvalue_left) && !dtvalue_left.Contains(";") && !String.IsNullOrWhiteSpace(dtvalue_right))
                        {
                            double pokleft = Convert.ToDouble(value_left);
                            double pokright = Convert.ToDouble(value_right);

                            // если не нарушен нарастающий итог
                            if (pokleft <= pokright)
                            {
                                System.TimeSpan deltaday = Convert.ToDateTime(dtvalue_right) - Convert.ToDateTime(dtvalue_left);
                                double deltapok = pokright - pokleft;

                                double srednesut_calc = deltapok / deltaday.Days;
                                double srmes_calc = Math.Round(srednesut_calc * DateTime.DaysInMonth(Convert.ToInt32(periodyeartek), Convert.ToInt32(periodmonthtek)));

                                // формируем отчет -----------------------------------------------

                                /*if (!String.IsNullOrWhiteSpace(srmes_iesbk_str)) // еслм СрМес ПО ИЭСБК присутствует
                                {
                                    DateTime dt_period = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);

                                    if (srmes_iesbk >= 0 && Convert.ToDateTime(dtvalue_right).CompareTo(dt_period) < 0) // не выводим наши расчеты, если СрМес ИЭСБК < 0 и правое показание принадлежит текущему периоду анализа
                                    {
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(srmes_calc); // СрМес РАСЧ                                

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Red;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(srmes_calc - srmes_iesbk); // Недополученный ПО                                
                                    }
                                }*/

                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue(srmes_calc);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;

                            } // if (pokleft <= pokright)
                              //-----------------------------------------------------------------
                        } // if (!String.IsNullOrWhiteSpace(dtvalue_left) && !String.IsNullOrWhiteSpace(dtvalue_right))

                    } // if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    // end расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)
                }

                splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + tableTOTALls10.Rows.Count.ToString() + ")");

            } // for (int i = 0; i < tableTOTALls10.Rows.Count; i++)

            // форматируем строку-заголовок
            worksheet.Rows[0].Font.Bold = true;
            worksheet.Rows[0].Alignment.WrapText = true;
            worksheet.Rows[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Rows[0].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            worksheet.Rows[0].AutoFit();

            worksheet.Columns.AutoFit(0, MAXCOLinWRKSH);

            worksheet.Columns.Group(3, 7, true); // группируем по колонкам "ФИО" - "Номер квартиры"            

            //int[] propidmas = new int[] { 1006, 1007, 1050, 1024, 1025, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };
            //int[] propidmas = new int[] { 1051, 1024, 1025, 1006, 1007, 1050, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };

            // группируем "периодические" значения - в частности расшифровку ПО
            for (int period_i = 0; period_i < MAX_PERIOD_MONTH; period_i++)
            {
                //worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period, FIRST_COLUMNS + period_i * columns_in_period + 2, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 2 + 1, FIRST_COLUMNS + period_i * columns_in_period + 4 + 1, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 7 + 1, FIRST_COLUMNS + period_i * columns_in_period + 7 + 10 + 1, true); // после "ПолОтп"
            }

            // группируем последние колонки анализа ПО по показаниям ПУ
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1 + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1, true);

            // группируем колонки слагаемых ПО от ИЭСБК
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 2 + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1 + MAX_PERIOD_MONTH + 1, true);

            worksheet.FreezeRows(0); // "фиксируем" верхнюю строку

            form1.spreadsheetControl1.EndUpdate();

            tableIESBKotdelenie.Dispose();
            tablelsPropGLOBAL.Dispose();

            SQLconnection.Close();
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // отчет-"шахматка" по наличию л/с и полезного отпуска - 2        
    }
}
