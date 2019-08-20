//Опишите свой класс и его методы ниже. Данная сборка будет доступна в namespace        


using System.Collections.Generic;
using System.Linq;
using ITnet2.Common.Tools;
using ITnet2.Server.BusinessLogic.Core.DocumentConfig;
using ITnet2.Server.BusinessLogic.Core.Documents;
using ITnet2.Server.BusinessLogic.Core.Documents.Forming;
using ITnet2.Server.Data;
using ITnet2.Server.Dialogs;
using ITnet2.Server.Session;
using ITnet2.Server.BusinessLogic.Core.Bindings;
using ITnet2.Server.BusinessLogic.LP.Schedules;
using System;
using System.Linq;
using ITnet2.Server.Data.Tables;
using ITnet2.Common.Tools.FileConverter;
using System.Text;
using System.Xml;
using ITnet2.Server.UserBusinessLogic._Setdromos;
using ITnet2.Server.UserBusinessLogic.Docworkflow;
using ITnet2.Server.UserBusinessLogic.Resources;

using Kum = ITnet2.Server.UserBusinessLogic.Kum;

using ITnet2.Server.Data;
using ITnet2.Server.BusinessLogic.LP.Accounting;
using ITnet2.Server.Session;
using System.Collections.Generic;
using ITnet2.Server.Data.Tables;
using ITnet2.Server.BusinessLogic.Scenario.Dependencies;
using ITnet2.Server.BusinessLogic.Core.Documents;
using ITnet2.Server.BusinessLogic.Core.Documents.Forming;
using ITnet2.Server.Dialogs;
using ITnet2.Server.BusinessLogic.LP.PublicSurface.Accounting;
using ITnet2.Server.BusinessLogic.Core.Analytics;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Linq.Expressions;
using System.Collections;



public class VMVPRO
{
    static int counter = 0;

    public static string tableName = "";

    private static bool _isTableExists = false;

    private string s;

    public void Run(string path, string fileName)
    {

        OpenExcel(path, fileName);
        //Test("011 - 13100 - 1100045 - Сіманенко - test");

    }

    public void OpenExcel(string path, string fileName)
    {
        OleDbConnection connection;

        try
        {
            if (!InfoManager.YesNo("OpenExcel_")) return;

            decimal undoc;
            string ndm_s;

            string BS = "00223";        // Забалонсовый счет

            string[] stringSeparator = new string[] { " - " };
            string[] result = fileName.Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);

            int cehDMZ = Convert.ToInt32(result[1].ToString());
            string n_kdkDMZ = result[2].ToString();

            string connectionString;

            //'Для Excel 12.0 
            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + fileName + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";
            connection = new OleDbConnection(connectionString);
            connection.Open();

            OleDbCommand command = connection.CreateCommand();

            command.CommandText = "Select * From [sheet$A0:I15000] "; //Where [К-во (факт)] = 20 ";

            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataTable dt = new DataTable();
            da.Fill(dt);

            //if (!InfoManager.YesNo("InsertBalanceDMZ")) return;
            if (!InsertBalanceDMZ(out undoc, out ndm_s, cehDMZ, n_kdkDMZ, fileName))
            {
                InfoManager.MessageBox("DMZ не вставилась!!!");
                return;
            }
            //if (!InfoManager.YesNo("Вставка прошла DMZ")) return;

            //if (!InfoManager.YesNo("Подготовка на вставку строк DMS")) return;

            RowTXT rowTXT = new RowTXT();
            rowTXT.CloneTable(dt);

            List<string> DoubleKmat = new List<string>();

            int k = 0;
            foreach (DataRow row in dt.Rows)
            {
                //if (!InfoManager.YesNo("Прдолжить добавление строк?")) return;

                try
                {
                    string ceh_s = row["ceh"].ToString();
                    string kmat_old = row["kmat"].ToString();

                    //string kmat = ConvertKmat(kmat_old, ceh_s);
                    string kmat = ConvertKmat(kmat_old, ceh_s, DoubleKmat);

                    bool flag1 = false;
                    object[] arrayColumn = row.ItemArray;

                    if (Flag(arrayColumn)) break;

                    int ceh = Convert.ToInt32(ceh_s);

                    string n_kdk = row["n_kdk"].ToString();
                    string naim = row["naim"].ToString();
                    string size_type = row["size_type"].ToString();
                    int ei = FuncEI(row["ei"].ToString());
                    decimal price = FuncPrice(row["price"].ToString());
                    decimal count = FuncCount(row["count"].ToString());
                    
                    //InfoManager.MessageBox("row['sum'].ToString() = " + row["sum"].ToString());
                    decimal sum = FuncSum(row["sum"].ToString());

                    #region "   ShowKmat   "

                    //StringBuilder sb = new StringBuilder();

                    //sb.Append(ceh.ToString() + "\n");
                    //sb.Append(kmat.ToString() + "\n");
                    //sb.Append(naim.ToString() + "\n");
                    //sb.Append(size_type.ToString() + "\n");
                    //sb.Append(ei.ToString() + "\n");
                    //sb.Append(price.ToString() + "\n");
                    //sb.Append(count.ToString() + "\n");
                    //sb.Append(sum.ToString() + "\n");

                    ////InfoManager.MessageBox(sb.ToString());

                    //if (!InfoManager.YesNo(sb.ToString())) return;

                    #endregion

                    //=================================================================

                    try
                    {
                        if (!DoubleKmat.Contains(kmat_old) | kmat_old == "")
                        {
                            DoubleKmat.Add(kmat_old);

                            DataRow[] rows1 = null;
                            try
                            {
                                rows1 = dt.Select("kmat = '" + kmat_old + "' and naim = '" + naim + "'");
                            }
                            catch (Exception)
                            {
                                rows1 = dt.Select("kmat = " + kmat_old + " and naim = '" + naim + "'");
                            }


                            int flag = 0;

                            if (rows1.Count() > 1)
                            {
                                flag = 1;
                                for (int i = 0; i < rows1.Count(); i++)
                                {
                                    DataRow r1 = rows1[i];

                                    naim = r1["naim"].ToString();
                                    size_type = r1["size_type"].ToString();
                                    ei = FuncEI(r1["ei"].ToString());
                                    price = FuncPrice(r1["price"].ToString());
                                    count = FuncCount(r1["count"].ToString());
                                    sum = FuncSum(r1["sum"].ToString());

                                    if (i == 0)
                                    {
                                        string ss1 = rows1[0].ItemArray[2].ToString();
                                        string ss2 = rows1[0].ItemArray[3].ToString();



                                        string sss = rows1.ToString();
                                        if (!KsmTable.IsRecord(kmat))
                                            InsertKmat(kmat, kmat_old, naim, size_type, ei, fileName, BS);

                                        try
                                        {
                                            InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, ei, count, price, sum, BS);
                                        }
                                        catch (Exception)
                                        {
                                            if (!InfoManager.YesNo("Insert False")) return;
                                        }



                                        //KSM.Add(kmat, ss2);
                                        //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));

                                    }
                                    else
                                    {
                                        try
                                        {
                                            InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, ei, count, price, sum, BS);
                                        }
                                        catch (Exception)
                                        {
                                            if (!InfoManager.YesNo("Insert False")) return;
                                        }


                                        //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                    }

                                }
                            }

                            DataRow[] rows2 = null;
                            try
                            {
                                rows2 = dt.Select("kmat = '" + kmat_old + "'");
                            }
                            catch (Exception)
                            {
                                rows2 = dt.Select("kmat = " + kmat_old);
                            }

                            if (flag == 0)
                            {
                                if (rows2.Count() > 1)
                                {
                                    for (int i = 0; i < rows2.Count(); i++)
                                    {
                                        DataRow r2 = rows2[i];

                                        naim = r2["naim"].ToString();
                                        size_type = r2["size_type"].ToString();
                                        ei = FuncEI(r2["ei"].ToString());
                                        price = FuncPrice(r2["price"].ToString());
                                        count = FuncCount(r2["count"].ToString());
                                        sum = FuncSum(r2["sum"].ToString());

                                        if (i == 0)
                                        {
                                            //naim = r2["naim"].ToString();
                                            //size_type = r2["size_type"].ToString();

                                            if (!KsmTable.IsRecord(kmat))
                                                InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                                            try
                                            {
                                                InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                            }
                                            catch (Exception)
                                            {
                                                if (!InfoManager.YesNo("Insert False")) return;
                                            }

                                            //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                            //}
                                        }
                                        else
                                        {

                                            counter++;
                                            kmat = ConvertKmat("", ceh_s, DoubleKmat);

                                            if (!KsmTable.IsRecord(kmat))
                                                InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                                            try
                                            {
                                                InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                            }
                                            catch (Exception)
                                            {
                                                if (!InfoManager.YesNo("Insert False")) return;
                                            }


                                            //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                        }

                                    }
                                }
                                else
                                {

                                    //string ss1 = rows2[0].ItemArray[2].ToString();
                                    //string ss2 = rows2[0].ItemArray[3].ToString();
                                    //if (!KsmTable.IsRecord(kmat))
                                    //{
                                    //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                                    if (kmat_old == "")
                                    {
                                        counter++;
                                        kmat = ConvertKmat("", ceh_s, DoubleKmat);
                                        //kmat = ConvertKmat("", ceh_s);
                                    }

                                    if (!KsmTable.IsRecord(kmat))
                                        InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                                    try
                                    {
                                        InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                    }
                                    catch (Exception)
                                    {
                                        if (!InfoManager.YesNo("Insert False")) return;
                                    }


                                    //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                }
                            }


                            //else
                            //    InfoManager.MessageBox(String.Format("Кода {0} нет!", kmat));

                            //KSM.Add(rowKmat["kmat"].ToString(), rowKmat["kmat"].ToString());
                            //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}", k, ceh, kmat, kmat_old, ei, price, count, sum));
                        }

                    }
                    catch (Exception)
                    {
                        InfoManager.MessageBox("Ошибка!");

                        //DoubleKmat.Add(kmat_old);

                        //var rowsKmat = dt.AsEnumerable().Where(x => x["kmat"].ToString() == kmat_old);
                        var rowsKmat = dt.Select("kmat = '" + kmat_old + "' ");

                        foreach (var rowKmat in rowsKmat)
                            rowTXT.Add(rowKmat, dt.Rows.IndexOf(rowKmat) + 2);

                    }

                    //=================================================================

                    //if (!InfoManager.YesNo("InsertKSM " + kmat)) return;
                    //if (!InfoManager.YesNo("InsertBalanceDMS")) return;


                }
                catch (Exception ex)
                {
                    //if (!InfoManager.YesNo(ex.Message)) return;

                    bool flag = false;
                    object[] ddd = row.ItemArray;
                    for (int i = 0; i < ddd.Length - 1; i++)
                    {
                        flag = ddd[i].ToString() == "" ? true : false;
                    }

                    if (flag) break;


                    rowTXT.Add(row, k);
                }
            }

            rowTXT.WriteTXT(path, fileName, n_kdkDMZ);

            try
            {
                connection.Close();
            }
            catch (Exception ex)
            {
                if (!InfoManager.YesNo("Ошибка:\n" + ex.Message)) return;
            }
        }
        catch (Exception ex)
        {
            if (!InfoManager.YesNo("Ошибка:\n" + ex.Message)) return;
        }

    }

    void ShowKmat(StringBuilder sb)
    {
        //sb.Append(ceh.ToString() + "\n");
        //sb.Append(kmat.ToString() + "\n");
        //sb.Append(naim.ToString() + "\n");
        //sb.Append(type.ToString() + "\n");
        //sb.Append(ei.ToString() + "\n");
        //sb.Append(price.ToString() + "\n");
        //sb.Append(count.ToString() + "\n");
        //sb.Append(sum.ToString() + "\n");

        //InfoManager.MessageBox(sb.ToString());
    }

    /// <summary>
    /// Функция для включения/выключения ТМЦ (Принимает сформированный лист и параметр enable = true/включить, false/выключить )
    /// </summary>
    /// <param name="?"></param>
    void EnableKmat(List<string> kmats, bool enable)
    {
        if (!InfoManager.YesNo("Вход EnableKmat(kmats)")) return;

        int k = 0;
        for (int i = 0; i < kmats.Count; i++)
        {
            var OldKmat = KsmTable.GetRecord(kmats[i]);

            var tablePOD_OLD = PodTable.GetRecord(OldKmat.Kskl);

            var upd = SqlClient.Main.CreateUpdateBuilder();

            upd.Table.Name = "KSM";

            if (enable)
                upd.Set.SetValue("PR_DO", "");      // Включить
            else
                upd.Set.SetValue("PR_DO", "О");     // Выключить

            upd.Where = new SqlCmdText("KSM.kmat=@kmat", new SqlParam("kmat", kmats[i]));

            upd.Exec();
        }
    }

    void NextStateAndPostingDo(int Undoc)
    {
        // Следуящая стадия
        //InfoManager.MessageBox("Процедура включения");

        DocumentWorkflow dwf = new DocumentWorkflow();
        DocumentWorkflow.MoveNextChain(Undoc);

        //DocumentWorkflow dwf = new DocumentWorkflow();
        //InfoManager.MessageBox("Переход на след. стадию произошел");
        //DocumentWorkflow.DoPosting(Undoc);

        //InfoManager.MessageBox("Разноска произошла успешно");
        //DocumentWorkflow.Refresh();
    }

    /// <summary>
    /// Поставить документ на бизнес-процесс
    /// </summary>
    /// <param name="undoc"></param>
    void FixStageBusinessProcess(decimal undoc)
    {
        var filter2 = new SqlCmdText("DMZ.UNDOC in (@UNDOC)", new SqlParam("UNDOC", undoc) { Array = true });

        var filter = new SqlCmdText(undoc.ToString());

        var si2 = new DataEditor.StartInfo("DMZ10")
        {
            Editable = true,
            StartMode = new DataEditor.StartInfo.DataEditorStartMode(new DataEditor.StartInfo.WorkflowStartMode(WorkflowProcessMode.SetRoute) { AllRecord = true, BatchMode = true }),
            SkipFilterDialogs = true,
            AdditionalFilter = filter2
        };

        DataEditor.Call(si2);
    }

    dynamic grid;

    public class RessorcePair
    {
        public string KMAT_OLD;
        public string KMAT_NEW;
    }

    void InsertBalanceDMS(decimal undoc, string ndm, int ceh, string n_kdk, string kmat, int edi, decimal kol, decimal cena, decimal sum, string db)
    {
        int cnt = 1;
        int npp = (SqlClient.CreateCommand(string.Format("Select Max(NPP) From DMS Where Undoc = @undoc "), new SqlParam("undoc", undoc)).ExecScalar<int>() + cnt);

        //DateTime DT = DateTime.Now;
        //DateTime DT = new DateTime(2017, 10, 31);
        DateTime DT = new DateTime(2019, 08, 1);

        var ib = SqlClient.Main.CreateInsertBuilder();

        ib.TableName = "DMS";

        ib.AddValue("KURS", 1);
        ib.AddValue("UNDOC", undoc);
        ib.AddValue("NPP", npp);
        ib.AddValue("GODMEC", "201811");
        ib.AddValue("ALIAS", "DMZ");
        ib.AddValue("DDM", DT);
        ib.AddValue("NDM", ndm);
        ib.AddValue("KDMT", "BALOBOR");
        ib.AddValue("KSTO", 0);
        ib.AddValue("KOBJ_CR", "CO000");
        ib.AddValue("KOBJ_DB", "CO000");
        ib.AddValue("CEH_K", ceh);
        ib.AddValue("N_KDK_K", n_kdk);
        ib.AddValue("KVAL", 980);
        ib.AddValue("KSTS", "001");
        ib.AddValue("KMAT", kmat);
        ib.AddValue("EDI", edi);
        ib.AddValue("KOL", kol);         // количество
        ib.AddValue("EDI2", edi);
        ib.AddValue("KOL_EDI2", kol);
        ib.AddValue("CENA_1", cena);           // Цена
        ib.AddValue("CENA_1VAL", cena);        // Цена
        ib.AddValue("SUMMA_1", sum);         // Сумма
        ib.AddValue("SUMMA_1VAL", sum);      // Сумма
        ib.AddValue("CENA_2", cena);          // Количество
        ib.AddValue("CENA_2VAL", cena);       // Количество
        ib.AddValue("SUMMA_2", sum);         // Сумма
        ib.AddValue("SUMMA_2VAL", sum);      // Сумма
        ib.AddValue("KBLS", "UKR");
        ib.AddValue("DB", db);
        ib.AddValue("ORG_DB", 1);
        ib.AddValue("KDOG_DB", "БЕЗ ДОГОВОРА");
        ib.AddValue("UNDOG_DB", 4531);
        ib.AddValue("UNDOG_CR ", 4531);

        ib.AddValue("KAU4DB", "НЕТ");
        ib.AddValue("KZAJ_DB", "НЕТ");
        ib.AddValue("KDOG_CR", "БЕЗ ДОГОВОРА");

        ib.AddValue("KSTO", 246);

        ib.AddValue("PR_P", "+");
        ib.AddValue("GM_MBP", "201811");
        ib.AddValue("FIO_D", "ADM");
        ib.AddValue("DATE_D", DateTime.Now);

        //if (!InfoManager.YesNo("Перед самой вставкой DMS")) return;
        ib.Exec();



    }

    bool InsertBalanceDMZ(out decimal undoc, out string ndm_s, int CEH, string N_KDK, string fileName)
    {
        bool res = false;

        undoc = 0;
        ndm_s = "";

        try
        {
            decimal cnt = 1;
            undoc = SqlClient.CreateCommand("Select LAST_NOM From LSTN Where Alias = 'DMR     ' ").ExecScalar<decimal>() + cnt;
            decimal ndm = SqlClient.CreateCommand(string.Format("Select LAST_NOM From LSTN Where ALIAS = 'DMZ' and LSTNOBJ = 'CO000CBALOBORM201712' ")).ExecScalar<decimal>() + cnt;

            ndm_s = ndm.ToString();

            //if (!InfoManager.YesNo("undoc = " + undoc.ToString() + "\n" + "ndm = " + ndm.ToString())) return;

            SqlClient.CreateCommand(string.Format("Update LSTN Set LAST_NOM = @undoc Where ALIAS = 'DMR     ' "), new SqlParam("undoc", undoc)).ExecNonQuery();
            SqlClient.CreateCommand(string.Format("Update LSTN Set LAST_NOM = @undoc Where ALIAS = 'DMZ' and LSTNOBJ = 'CO000CBALOBORM201712' "), new SqlParam("undoc", ndm)).ExecNonQuery();

            //
            //DateTime DT = DateTime.Now;
            //DateTime DT = new DateTime(2017, 10, 31);

            // DateTime DT = new DateTime(2018, 11, 1);

            DateTime DT = new DateTime(2019, 08, 1);


            var ib = SqlClient.Main.CreateInsertBuilder();

            ib.TableName = "DMZ";

            ib.AddValue("UNDOC", undoc);
            ib.AddValue("GODMEC", "201811");
            ib.AddValue("DDM", DT);
            ib.AddValue("NDM", ndm_s.ToString());
            //ib.AddValue("KDMT", "_BALANCES");
            ib.AddValue("KDMT", "BALOBOR");
            ib.AddValue("KSTO", 0);

            ib.AddValue("KOBJ_CR", "CO000");
            ib.AddValue("KOBJ_DB", "CO000");

            ib.AddValue("CEH_K", CEH);
            ib.AddValue("N_KDK_K", N_KDK);      // Получатель               - N_KDK_UMC

            ib.AddValue("ORG_2", 1);
            ib.AddValue("NBNK2", 10);
            ib.AddValue("ORG_GPO_2", 1);

            ib.AddValue("KVAL", 980);
            ib.AddValue("KKVT", "NB");
            ib.AddValue("KURS", 1);
            ib.AddValue("PRRA", "+");
            ib.AddValue("KSD2", "05");

            ib.AddValue("KSTO", 246);

            ib.AddValue("COMM", fileName);

            ib.AddValue("KBLS", "UKR");
            ib.AddValue("FIO_D", "ADM");
            ib.AddValue("DATE_D", DateTime.Now);

            ib.Exec();

            FixStageBusinessProcess(undoc);

            res = true;

            return res;
        }
        catch (Exception ex)
        {

            res = false;
            InfoManager.MessageBox("False DMZ");

            return res;
        }


    }

    public class MyRecord
    {
        public int Edi;
        public string Nedi;
    }

    public void TestKSM2()
    {
        int Undoc = 405193;

        DocumentWorkflow dwf = new DocumentWorkflow();

        // Пполучение документа внутри класса
        DocumentWorkflow.HeadDocumentInChain(Undoc);
        //InfoManager.MessageBox("DocumentWorkflow.HeadDocumentInChain(Undoc);");

        // Переход на стадию
        DocumentWorkflow.MoveNextChain(Undoc);
        //InfoManager.MessageBox("DocumentWorkflow.MoveNextChain(Undoc);");

        //DocumentWorkflow.MovePreviousStageIncome();
        //InfoManager.MessageBox("DocumentWorkflow.MovePreviousStageIncome();");

        // Пример проведения проводок по документу спомощью RRT метода
        DocumentWorkflow.CallRRTMetod(Undoc, "DMZ_DMR10");
        //InfoManager.MessageBox("DocumentWorkflow.CallRRTMetod(Undoc, 'DMZ_DMR10');");

        // Пример проведения проводок по документу
        DocumentWorkflow.FormEntries(Undoc);  //FormEntries(int Undoc)
        //InfoManager.MessageBox("DocumentWorkflow.FormEntries(Undoc)");

        // Пример удалления проводок по документу
        DocumentWorkflow.DeleteEntries(Undoc);
        //InfoManager.MessageBox("DocumentWorkflow.DeleteEntries(Undoc);");



        DocumentWorkflow.Refresh();
        //InfoManager.MessageBox("DocumentWorkflow.Refresh();");

    }


    public void InsertKmat(string kmat, string old_kmat, string nmat, string naimkm_s, int ei, string fileName, string BS)
    {
        //if (!InfoManager.YesNo("До вставки Шифра")) return;

        //if (KsmTable.IsRecord(kmat))
        //{
        //    InfoManager.MessageBox(String.Format("Код {0} есть!", kmat));
        //    return;
        //}
        //else
        //    InfoManager.MessageBox(String.Format("Кода {0} нет!", kmat));

        //return;

        var ib = SqlClient.Main.CreateInsertBuilder();

        //string kmat = "917000000000004";

        //if (KsmTable.IsRecord(old_kmat))
        //    InfoManager.MessageBox(String.Format("Код {0} есть!", old_kmat));
        //else
        //    InfoManager.MessageBox(String.Format("Кода {0} нет!", old_kmat));

        //return;

        ib.TableName = "KSM";

        ib.AddValue("KMAT", kmat);
        ib.AddValue("NMAT", nmat);
        ib.AddValue("NAIMKM_S", naimkm_s);
        ib.AddValue("N_RES", nmat + " " + naimkm_s);
        ib.AddValue("N_RES_DOC", nmat + " " + naimkm_s);
        ib.AddValue("PRNAIM", 0);
        ib.AddValue("SKM", "920");                      //SKM	917	Група, підгрупа ресурсів	C	15

        ib.AddValue("KKST", "D");                       //KKST	D	Тип ресурсу	C	1
        ib.AddValue("EDI", ei);                        //EDI	796	Основна одиниця виміру (облікова). Код	N	3
        ib.AddValue("EDI_NORM", ei);                   //EDI_NORM	796	ОВ норм для матеріалів. Код	N	3
        ib.AddValue("EDI_NORMP", ei);                  //EDI_NORMP	796	ОВ норм для продукції. Код	N	3
        ib.AddValue("KPER_NORM", 1);                    //KPER_NORM	1.000000000	Коефіцієнт переведення ОВ->ОВ в нормах	N	16
        ib.AddValue("KPER_N_I", 1);                     //KPER_N_I	1.000000000	Коефіцієнт переведення ОВ в нормах->ОВ	N	16
        ib.AddValue("EDI2", ei);                       //EDI2	796	Одиниця виміру-2. Код	N	3
        ib.AddValue("KPER2", 1);                        //KPER2	1.000000000	Коефіцієнт переведення ОВ->ОВ2	N	16
        ib.AddValue("KPER2_I", 1);                      //KPER2_I	1.000000000	Коефіцієнт переведення ОВ2>ОВ	N	16
        ib.AddValue("EDI3", ei);                       //EDI3	796	Одиниця  виміру - 3. Код	N	3
        ib.AddValue("KPER3", 1);                        //KPER3	1.000000000	Коефіцієнт переведення ОВ->ОВ3	N	16
        ib.AddValue("KPER3_I", 1);                      //KPER3_I	1.000000000	Коефіцієнт переведення ОВ3>ОВ	N	16
        ib.AddValue("EDI_NKALK", ei);                  //EDI_NKALK	796	ОВ нормативних калькуляцій. Код	N	3
        ib.AddValue("GOST", "");                   //GOST	ГОСТ	ДСТ, ТУ	C	60
        ib.AddValue("KBLS", "UKR");                     //KBLS	UKR	План рахунків. Код	C	5
        ib.AddValue("BS", BS);                      //BS	0022	Балансовй рахунок зберігання на складах (основний)	C	10
        ib.AddValue("OLDKMAT", old_kmat);    //OLDKMAT	Старый код 00000	Ресурс "старий". Код	C	30
        ib.AddValue("COMM", "файл Excel " + fileName);             //COMM	Примечание	Примiтка	C	254
        ib.AddValue("MARKA_", "");               //MARKA_	Марка	Марка	C	100

        ib.Exec();

        //InfoManager.MessageBox("Вставка прошла успешно!");

    }

    static public bool Flag(object[] arrayColumn)
    {
        bool result2 = false;
        for (int i = 0; i < arrayColumn.Length - 1; i++)
        {
            if (arrayColumn[i].ToString() == "")
            {
                result2 = true;
            }
            else
                return false;
        }

        return result2;
    }

    static public int FuncEI(string ei)
    {
        try
        {
            if (ei == "" || ei == "796")
                return ConvertEI1(ei); //Convert.ToInt32("796");
            else
            {
                return ConvertEI2(ei);
            }
        }
        catch (Exception)
        {
            return 796;
        }
    }

    static public int ConvertEI1(string ei)
    {
        
        return Convert.ToInt32(ei);

        
    }

    static public int ConvertEI2(string ei)
    {

        //switch (ei)
        //{
        //    case "1": return 839;
        //    case "2": return 796;
        //    case "3": return 166;
        //    case "4": return 163;
        //    case "5": return 6;
        //    case "6": return 761;
        //    case "7": return 168;
        //    case "8": return 798;
        //    case "9": return 797;
        //    case "10": return 112;
        //    case "11": return 736;
        //    case "796": return 796;
        //}
        
        return Convert.ToInt32(ei);

        //return 796;
    }

    static public decimal FuncPrice(string price_)
    {
        decimal price = 0;
        try
        {
            string priceString = price_.Replace('.', ',');
            price = Convert.ToDecimal(priceString);
        }
        catch (Exception)
        {
            price = 0;
        }

        return price;
    }

    static public decimal FuncSum(string sum_)
    {
        decimal sum = 0;
        try
        {
            string sumString = sum_.Replace('.', ',');
            sum = Convert.ToDecimal(sumString);
            
        }
        catch (Exception ex)
        {
            //InfoManager.MessageBox(ex.Message);
            sum = 0;
        }

        return sum;
    }

    static public decimal FuncCount(string count_)
    {
        decimal count = 0;
        try
        {
            string countString = count_.Replace('.', ',');
            count = Convert.ToDecimal(countString);
        }
        catch (Exception)
        {
            count = 0;
        }

        return count;
    }

    static string ConvertKmat(string kmat_old, string ceh, List<string> DoubleKmat)
    {
        string kmat = "";
        string ceh_convert = "";
        int count_kmat_old = 0;

        string old_kmat_str = "";
        try
        {
            string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
            old_kmat_str = Convert.ToInt32(old_kmat_convert).ToString();   // 8
        }
        catch (Exception)
        {
            string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
            old_kmat_str = old_kmat_convert;
        }

        if (!DoubleKmat.Contains(kmat_old) || kmat_old == "")
        {

        }

        if (ceh.Count() < 6 && old_kmat_str.Count() <= 7)
            ceh_convert = ceh;
        else if (ceh.Count() > 4)
            ceh_convert = ceh.ToString().Substring(0, 1) + ceh.ToString().Substring(2, 3);
        else
            ceh_convert = ceh;

        int len = ceh.Count();

        if (kmat_old == "" | DoubleKmat.Contains(kmat_old))
        {
            //counter++;
            string str_counter = counter.ToString();
            int len_counter = str_counter.Length;
            int len_ceh = ceh.Length;
            int symbols = 11 - len_ceh - len_counter;
            len_counter = 0;
            //if (str_counter.Length == 1) len_counter = 

            //kmat = "917" + "vmv" + ceh + new String('0', 15 - 2 - str_counter.Length -  len_ceh - symbols) + str_counter;
            kmat = "920" + "vvv" + ceh + new String('0', 5 - str_counter.Length) + str_counter;


            return kmat;
        }

        int len_kmat_old = old_kmat_str.Count();
        
        if (len_kmat_old >= 12 && !DoubleKmat.Contains(kmat_old))
        {
            kmat = "920" + old_kmat_str.Substring(len_kmat_old - 12, 12);
        }
        else if (old_kmat_str.Count() == 11)
        {
            kmat = "920" + "0" + old_kmat_str;
        }
        else if (old_kmat_str.Count() == 10)
        {
            kmat = "920" + "00" + old_kmat_str;
        }
        else if (old_kmat_str.Count() == 9)
        {
            kmat = "920" + ceh.Substring(len - 3, 3) + old_kmat_str;
        }
        else if (old_kmat_str.Count() == 8)
        {
            kmat = "920" + ceh_convert + old_kmat_str;
        }
        else
        {
            count_kmat_old = 12 - ceh_convert.ToString().Count() - old_kmat_str.Count();
            kmat = "920" + ceh_convert.ToString() + new String('0', count_kmat_old) + old_kmat_str;   // 3 + 4 + 1 + 7
        }

        return kmat;
    }

    static string ConvertKmat_(string kmat_old, string ceh)
    {
        string kmat = "";
        string ceh_convert = "";
        int count_kmat_old = 0;

        string old_kmat_str = "";
        try
        {
            string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
            old_kmat_str = Convert.ToInt32(old_kmat_convert).ToString();   // 8
        }
        catch (Exception)
        {
            string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
            old_kmat_str = old_kmat_convert;
        }


        if (ceh.Count() < 6 && old_kmat_str.Count() <= 7)
            ceh_convert = ceh;
        else if (ceh.Count() > 4)
            ceh_convert = ceh.ToString().Substring(0, 1) + ceh.ToString().Substring(2, 3);
        else
            ceh_convert = ceh;

        int len = ceh.Count();

        if (kmat_old == "")
        {
            //counter++;
            string str_counter = counter.ToString();
            int len_counter = str_counter.Length;
            int len_ceh = ceh.Length;
            int symbols = 11 - len_ceh - len_counter;
            len_counter = 0;
            //if (str_counter.Length == 1) len_counter = 

            //kmat = "917" + "vmv" + ceh + new String('0', 15 - 2 - str_counter.Length -  len_ceh - symbols) + str_counter;
            kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

            return kmat;
        }

        int len_kmat_old = old_kmat_str.Count();

        if (len_kmat_old > 15)
            kmat = old_kmat_str.ToString().Substring(1,15);
        else if (len_kmat_old >= 12)
            kmat = "920" + old_kmat_str.Substring(len_kmat_old - 12, 12);
        else if (old_kmat_str.Count() == 11)
            kmat = "920" + "0" + old_kmat_str;
        else if (old_kmat_str.Count() == 10)
            kmat = "920" + "00" + old_kmat_str;
        else if (old_kmat_str.Count() == 9)
            kmat = "920" + ceh.Substring(len - 3, 3) + old_kmat_str;
        else if (old_kmat_str.Count() == 8)
            kmat = "920" + ceh_convert + old_kmat_str;
        else
        {
            count_kmat_old = 12 - ceh_convert.ToString().Count() - old_kmat_str.Count();
            kmat = "920" + ceh_convert.ToString() + new String('0', count_kmat_old) + old_kmat_str;   // 3 + 4 + 1 + 7
        }

        return kmat;
    }

    public class RowTXT
    {
        DataTable dt;
        StringBuilder sb;
        string head = "";

        public void CloneTable(DataTable dt)
        {
            this.dt = dt.Clone();
        }

        public RowTXT()
        {
            dt = new DataTable("LogError");
            sb = new StringBuilder();

            head = String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}",
                                        "k".ToString().PadRight(5, ' '),
                                        "ceh".ToString().PadRight(6, ' '),
                                        "n_kdk".ToString().PadRight(8, ' '),
                                        "kmat_old".ToString().PadRight(12, ' '),
                                        "naim".ToString().PadRight(50, ' '),
                                        "size_type".ToString().PadRight(20, ' '),
                                        "ei".ToString().PadRight(9, ' '),
                                        "price".ToString().PadRight(9, ' '),
                                        "count".ToString().PadRight(8, ' '),
                                        "sum".ToString().PadRight(8, ' '));
        }

        public void Add(DataRow row, int k)
        {
            string ceh = row["ceh"].ToString();
            string n_kdk = row["n_kdk"].ToString();
            string kmat_old = row["kmat"].ToString() == "" ? "-------" : row["kmat"].ToString();
            string naim = row["naim"].ToString();
            string size_type = row["size_type"].ToString();

            string ei = row["ei"].ToString(); // == "" ? "-------" : row["ei"].ToString();
            string ei_temp = ei == "" || ei == "796" ? ConvertEI1(ei).ToString() : ConvertEI2(ei).ToString();

            if (ei_temp == "")
                ei = "-------";
            else if (ei_temp == "0")
                ei = ei + "(нет)";

            string price = row["price"].ToString() == "" ? "-------" : row["price"].ToString();
            string count = row["count"].ToString() == "" ? "-------" : row["count"].ToString();
            string sum = row["sum"].ToString();

            sb.AppendLine(String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}",
                                        k.ToString().PadRight(5, ' '),
                                        ceh.ToString().PadRight(6, ' '),
                                        n_kdk.ToString().PadRight(8, ' '),
                                        kmat_old.ToString().PadRight(12, ' '),
                                        naim.ToString().PadRight(50, ' '),
                                        size_type.ToString().PadRight(20, ' '),
                                        ei.ToString().PadRight(9, ' '),
                                        price.ToString().PadRight(9, ' '),
                                        count.ToString().PadRight(8, ' '),
                                        sum.ToString().PadRight(8, ' ')));

            dt.ImportRow(row);
        }



        public void WriteTXT(string path, string fileName, string n_kdk)
        {

            //decimal positionCount = dt.Rows.Count(  AsEnumerable().Count(x => x["n_kdk"].ToString() == n_kdk);
            decimal positionCount = dt.Select("n_kdk = '" + n_kdk.ToString() + "'").Count();
            decimal sumCount = 0;
            string sumCountString = "";



            try
            {
                //sumCount = dt.AsEnumerable().Sum(x => Convert.ToDecimal(x["count"]));

                sumCount = dt.Select("n_kdk = '" + n_kdk.ToString() + "'").Sum(x => Convert.ToDecimal(x["count"]));
                sumCountString = sumCount.ToString();
            }
            catch (Exception)
            {
                sumCountString = "Неопределено";
            }

            StringBuilder sb2 = new StringBuilder();
            if (sb.Length != 0)
            {
                sb2.AppendLine(head);
                sb2.AppendLine(sb.ToString());
                //sb2.Append("Количество номенклатуры, которая не попала:\t\" + sumCount);
                sb2.Append("К-во: " + CountPosition(positionCount) + new String('\t', 10) + "Всего: " + sumCountString);

                File.WriteAllText(path + fileName + ".txt", sb2.ToString());
            }
        }

        public string CountPosition(decimal count)
        {
            string[] text = new[] { " позиция", " позиций", " позиции" };

            string result = "";

            int len = count.ToString().Length;
            string endSymbol = count.ToString().Substring(len - 1, 1);

            if (count == 1)
                return count.ToString() + text[0];
            else if (count == 11 || count == 12 || count == 13 || count == 14)
                return count.ToString() + text[1];

            switch (endSymbol)
            {
                case "0":
                case "5":
                case "6":
                case "7":
                case "8":
                case "9":
                    result = text[1];
                    break;
                case "2":
                case "3":
                case "4":
                    result = text[2];
                    break;
                case "1":
                    result = text[0];
                    break;
            }

            return count.ToString() + result;
        }
    }

    public void Test(string fileName)
    {
        if (!InfoManager.YesNo("Test")) return;

        //string path = @"\\erp\TEMP\App\Остатки\";

        string path = @"\\erp\TEMP\App\Остатки\20.12.17\";
        //string path = @"\\erp\TEMP\App\Остатки\540\";
        //string path = @"\\erp\TEMP\App\Остатки\17.02.18\";


        // 

        //string fileName = "010 - 13004 - 1120665 - Лопатіна д";
        //string extension = "xlsx";

        string connectionString;
        OleDbConnection connection;
        //connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + fileName + "." + extension + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + fileName + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";
        connection = new OleDbConnection(connectionString);
        connection.Open();

        OleDbCommand command = connection.CreateCommand();

        command.CommandText = "Select * From [sheet$A0:I15000] "; //Where [К-во (факт)] = 20 ";

        OleDbDataAdapter da = new OleDbDataAdapter(command);
        DataTable dt = new DataTable();
        da.Fill(dt);

        //int s = dt.AsEnumerable().Count();
        try
        {
            var rows2 = dt.Select().AsEnumerable().Select(x => x["kmat"].ToString() == "001-100012").Single();

            InfoManager.MessageBox(rows2.ToString());

            //IEnumerable<DataTable> dd = (IEnumerable<DataTable>)dt;
            //System.Data.EnumerableRowCollection<System.Data.DataRow> dd = dt;

            //System.Data.EnumerableRowCollection<System.Data.DataRow> dd;
            //DataRow rowKmat = dt.AsEnumerable().Single(x => x["kmat"].ToString() == kmat_old);
            var rows = dt.Select("kmat = '001-100006' ");

            StringBuilder sb = new StringBuilder();
            //ceh	n_kdk	kmat	naim	size_type	ei	price	count	sum
            sb.AppendLine("Строк: " + rows.Count());

            foreach (DataRow row in rows)
            {
                sb.AppendLine("kmat = " + row["kmat"].ToString());
                sb.AppendLine("naim = " + row["naim"].ToString());
            }

            InfoManager.MessageBox(sb.ToString());

            InfoManager.MessageBox("Ok");
        }
        catch (Exception)
        {
            InfoManager.MessageBox("No");
        }

        //var objects =  

        connection.Close();

    }
}