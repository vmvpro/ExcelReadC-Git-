using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadC
{
    public class Program
    {
        static int counter = 0;

        static OleDbDataAdapter da;
        static DataTable dt;

        static void Main()
        {
            //dynamic app = null;
            //string excelProgId = "Excel.Application";
            //app = Activator.CreateInstance(Type.GetTypeFromProgID(excelProgId));

            //app.Visible = true;

            //app.Workbooks.Add();

            //return;

            //============================================================================

            //string fileName = "000 - 13012 - 1012103 - Test_1";

            //string fileName = "001 - 13006 - 1901076 - Шута м";
            //string fileName = "001 - 13006 - 1901076 - Шута м";

            //string fileName = "002 - 13006 - 1901076 - Шута д";

            //string fileName = "003 - 13011 - 1011105 - Тарасенко д_";
            //string fileName = "004 - 13011 - 1011105 - Тарасенко м";


            //string fileName = "005 - 13114 - 1114095 - Сімак  м";
            //string fileName = "006 - 13114 - 1114095 - Сімак  д";

            //string fileName = "007 - 13006 - 1006307 - Хомюк  м";
            //string fileName = "008 - 13006 - 1006307 - Хомюк  д";

            //string fileName = "009 - 13004 - 1120665 - Лопатіна м";
            //string fileName = "010 - 13004 - 1120665 - Лопатіна д";

            //string fileName = "011 - 13100 - 1100045 - Сіманенко.xls";
            //string fileName = "012 - 13533 - 1533309 - Фатюк";

            //string fileName = "013 - 13017 - 1017214 - Веску__.doc";
            //string fileName = "014 - 13003 - 1003008 - Кравченко  м__";

            //string fileName = "015 - 1003008 - Кравченко  д__";

            //string fileName = "016 - 13023 - 1023006 - Затенко__";

            //string fileName = "017 - 13001 - 1001024 - Пересенчук  м_1";
            //string fileName = "018 - 13001 - 1001024 - Пересенчук  д__";

            //string fileName = "019 - 13007 - 1007160 - Івасечко  д__";
            //string fileName = "020 - 13007 - 1007160 - Івасечко  м__";

            //string fileName = "022 - 13149 - 1149117 - Цимбалюк д_";
            //string fileName = "023 - 13149 - 1149117 - Цимбалюк м";

            //string fileName = "024 - 13019 - 1019138 - Вісленко  м";
            //string fileName = "025 - 13019 - 1019138 - Вісленко  д";

            //string fileName = "026 - 13021 - 1021078 - Матвєєва";
            //string fileName = "027 - 13015 - 1015004 - Губанова  д";    
            //string fileName = "028 - 13015 - 1015004 - Губанова  м";

            //-----------------------------------------------------------------------------

            //string fileName = "001 - 23016 - 1016156 - Черныш Л.П. - Балан_.xls";
            //string fileName = "002 - 23008 - 1008094 - Довгопол Т.Ф. - Балан_.xls";
            //string fileName = "003 - 23010 - 1010134 - Ковынева Т.Н. - Балан_.xls";

            //string fileName = "004 - 23045 - 1045084 - Панферова Т.Л. - Бондаренко_.xlsx";
            //string fileName = "005 - 23045 - 1045084 - Панферова Т. Л. - Бондаренко_.xlsx";

            //string fileName = "006 - 23015 - 1015282 - ПАНАСЮК І.Б. - Василенко_.xls";

            //string fileName = "007 - 23016 - 1016156 - Черныш Л.П. - Горбаченко_.xlsx";

            //string fileName = "008 - 23014 - 1016156 - Черныш Л.П. - Горбаченко_.xlsx";

            //string fileName = "009 - 23016 - 1016156 - Черныш Л.П. - Грищенко_.xls";

            //string fileName = "010 - 23008 - 1008094 - Довгопол Т.Ф. - Грищенко_.xls";
            //string fileName = "011 - 23010 - 1010134 - Ковынева Т.Н. - Грищенко_.xls";
            //string fileName = "012 - 23015 - 1015282 - ПАНАСЮК И.Б. - Петренко_.xlsx";

            //string fileName = "013 - 23015 - 1015282 - Панасюк - Середа_.xlsx";
            //string fileName = "014 - 23015 - 1015282 - Панасюк - Трофименкова_.xls";

            //string fileName = "015 - 23008 - 1008094 - Довгопол - Урбан_.xls";
            //string fileName = "016 - 23008 - 1008094 - Довгопол - Урбан_.xls";

            //string fileName = "017 - 23014 - 1016156 - Черниш Л.П. - Шияненко_.xlsx";
            //string fileName = "018 - 23014 - 1016156 - Черниш Л.П. - Шияненко_.xlsx";

            string path = @"\\erp\TEMP\App\Остатки\ЛИиДБ\";

            string fileName = @"
            003 - 13221 - 1221270 - Ковтун _vmv
            ";

            //DateTime dt = new DateTime(18,2,27);


            Main2(path.Trim(), fileName.Trim() + ".xlsx");
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
                if (ei == "")
                    return Convert.ToInt32("796");
                else
                {
                    return ConvertEI(ei);
                }
            }
            catch (Exception)
            {
                return 796;
            }
        }

        public static int ConvertEI(string ei)
        {
            switch (ei)
            {
                case "1": return 839;
                case "2": return 796;
                case "3": return 166;
                case "4": return 163;
                case "5": return 6;
                case "6": return 761;
                case "7": return 168;
                case "8": return 798;
                case "9": return 797;
                case "10": return 112;
                case "11": return 736;
                case "796": return 796;
            }

            return 796;
        }

        public static decimal FuncPrice(string price_)
        {
            decimal price = 0;
            try
            {
                price = Convert.ToDecimal(price_);
            }
            catch (Exception)
            {
                price = 0;
            }

            return price;
        }

        public static decimal FuncSum(string sum_)
        {
            decimal sum = 0;
            try
            {
                sum = Convert.ToDecimal(sum_);
            }
            catch (Exception)
            {
                sum = 0;
            }

            return sum;
        }

        public static decimal FuncCount(string count_)
        {
            decimal count = 0;
            try
            {
                count = Convert.ToDecimal(count_);
            }
            catch (Exception)
            {
                count = 0;
            }

            return count;
        }

        public static string ConvertKmat(string kmat_old, string ceh, List<string> DoubleKmat)
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

            if (kmat_old == "" || DoubleKmat.Contains(kmat_old))
            {
                string str_counter = counter.ToString();
                int len_counter = str_counter.Length;
                int len_ceh = ceh.Length;

                kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

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

        public static string ConvertKmatTest(string kmat_old, string ceh, List<string> DoubleKmat)
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

            //--------------------------------------------------------

            if (kmat_old == "" || DoubleKmat.Contains(kmat_old))
            {
                return CreateNewKmat(ceh, counter);
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

        private static string CreateNewKmat(string ceh, int counter)
        {
            string str_counter = counter.ToString();
            int len_counter = str_counter.Length;
            int len_ceh = ceh.Length;

            string kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

            return kmat;


        }

        public static Dictionary<string, string> KSM = new Dictionary<string, string>();

		static void  GroupByListRecource()
		{
			string old_kmat1 = "00123";
			string old_kmat2 = "00123";
			string old_kmat3 = "00222";
			string old_kmat4 = "00333";
			string old_kmat5 = "00333";

			var list = new List<string>() { "00123", "00123", "00222", "00333", "00333" };
			var listGroupBy = list.GroupBy(x => x);

			Dictionary<string, int> dic = new Dictionary<string, int>();

			foreach (var grp in listGroupBy)
			{
				dic.Add(grp.Key, grp.Count());
			}

			int key = listGroupBy.Where(x => x.Key == "00123").Count();

			//Dictionary<string, > dic_ = listGroupBy.ToDictionary();



			var listGroups =
				from prod in list
				group prod by list into prodGroup
				select new { OldKmat = prodGroup.Key, Count = prodGroup.Count() };


			//Dictionary<string, int> dic = new Dictionary<string, int>();

			//foreach (var item in listGroups)
			//	Console.WriteLine(item.OldKmat + " - " + item.Count);
			//    dic.Add(item.OldKmat, item.Count);



			//var where = listGroupBy.Where((x,y) => x == y);
		}

        static void Main2(string path, string fileName)
        {
            //string path = "Data.xlsx";
            //string path = "DataX.xlsx";

            DateTime DT = new DateTime(2017, 11, 1);

            // System.Data.EnumerableRowCollection<System.Data.DataRow> dd = dt;
            //System.Data.EnumerableRowCollection<System.Data.DataRow> dd;

            //string path = @"\\erp\TEMP\App\Остатки\";
            //string path = @"\\erp\TEMP\App\Остатки\20.12.17\";





            //string path = @"d:\Vetal\Work\MS Visual Studio\1_MyApplication\ExcelReadC\МШП передача (СЗА)\МШП_\";
            //string path = @"\\erp\TEMP\App\Остатки\540\";

            //string path = @"\\erp\TEMP\App\Остатки\17.02.18\";



            //-------------------------------------------------------

            string str_counter = "9999";
            int len_counter = str_counter.Length;
            string ceh99 = "13006";
            int len_ceh = ceh99.Length;
            int symbols = 11 - len_ceh - len_counter;
            len_counter = 0;
            //if (str_counter.Length == 1) len_counter =

            string kmat99 = "917" + "vmv" + ceh99 + new String('0', 4 - str_counter.Length) + str_counter;




            return;

            //string path = @"d:\Doc\Work\MS Visual Studio\ExcelReadC\Остатки\";

            //string fileName = "13100 - 1100045 - Сіманенко - vmv";
            //string fileName = "13100 - 1100045 - test";

            //string fileName = "000 - 13012 - 1012103 - Test_1";
            string extension = "xlsx";

            string[] stringSeparator = new string[] { " - " };
            string[] result;

            //917 - 1100 123-456-789

            result = fileName.Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);

            string ceh_ = result[1].Substring(0, 1) + result[1].Substring(2, 3);
            string n_kdk_file = result[2];

            //string kmat_s = "123-456-789";

            string kmat_s2 = "1234567890123";
            //string kmat_s2 = "1234567890123456";
            //string kmat_s2 = "1234-5678-9012-34";
            //string kmat_s2 = "1234-5678-9012-3";
            //string kmat_s2 = "1234-5678-9012";
            //string kmat_s2 = "1234-5678-9";
            string kmat_replace = kmat_s2.Replace("-", "");
            int len_kmat_s2 = kmat_replace.Count();

            //string kmat_ = ConvertKmat(kmat_replace, "13100");

            //if (kmat_replace.Count() >= 15)
            //    kmat_ = "917" + kmat_replace.Substring(len_kmat_s2 - 12, 12);
            //if (kmat_replace.Count() == 14)
            //    kmat_ = "917" + kmat_replace.Substring(len_kmat_s2 - 12, 12);
            //if (kmat_replace.Count() >= 12)
            //    kmat_ = "917" + kmat_replace.Substring(len_kmat_s2 - 12, 12);
            //if (kmat_replace.Count() == 11)
            //    kmat_ = "917" + "0" + kmat_replace.Substring(len_kmat_s2 - 11, 11);
            //if (kmat_replace.Count() == 10)
            //    kmat_ = "917" + "0" + kmat_replace.Substring(len_kmat_s2 - 11, 11);
            //if (kmat_replace.Count() == 9)
            //    kmat_ = "917" + "0" + kmat_replace.Substring(len_kmat_s2 - 11, 11);

            //return;
            //string kmat_old_ = Convert.ToInt32(kmat_s.Replace("-", "")).ToString();
            //string kmat_old2 = Convert.ToInt32("023-456-789".Replace("-", "")).ToString();

            //string kmat_new = "";
            //int count_kmat_old = 0;

            //if (kmat_old.Count() == 9)
            //    kmat_new = "917" + ceh.Substring(0,3) + kmat_old;
            //else if (kmat_old.Count() == 8)
            //    kmat_new = "917" + ceh + kmat_old;
            //else
            //{
            //    count_kmat_old = 12 - ceh.Count() - kmat_old.Count();
            //    kmat_new = "917" + ceh.ToString() + new String('0', count_kmat_old) + kmat_old;
            //}

            //string kmatTest = ConvertKmat("ДД01 23456700", "13100", DoubleKmat);

            //int int_kmat_old = kmat_old.Count();    // 7
            // 3 + 4 + 1 + 7
            // 3 + 3 + 0 + 9
            // 3 + 4 + 0 + 8



            //count_kmat_old = 12 - ceh.Count() - kmat_old.Count();
            //kmat_new = "917" + ceh.ToString() + new String('0', count_kmat_old) + kmat_old;

            //string kmat_new = kmat_old.Replace("-", "");



            string connectionString;
            OleDbConnection connection;


            //return;

            //'Для Excel 12.0 
            //connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + fileName + "." + extension + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";

            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + fileName + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";
            connection = new OleDbConnection(connectionString);
            connection.Open();

            OleDbCommand command = connection.CreateCommand();

            //command.CommandText = "Select * From [sheet$A9:D150]";
            //command.CommandText = "Select * From [Лист1$A6:G15000] "; //Where [К-во (факт)] = 20 "
            command.CommandText = "Select * From [sheet$A0:I15000] "; //Where [К-во (факт)] = 20 ";

            da = new OleDbDataAdapter(command);
            dt = new DataTable();

            da.Fill(dt);

            string name = dt.Rows[0]["n_kdk"].ToString();

            //var rows = dt.AsEnumerable().Where(x => x["kmat"].ToString() == "001-100006");
            //var rows = dt.AsEnumerable().Where(x => x.Field<string>("kmat") == "001-100006");

            //var rows999 = dt.Select("kmat ='" + "qwerty" + "' and naim = 'wwww' and price = 5").AsEnumerable();

            //string gkjhgk = (DataRow[])rows999[0][""];

            //var linq = from row1 in rows
            //           select new
            //           {
            //               kmat = row1["kmat"],
            //               naim = row1["naim"],
            //               price = row1["price"]
            //           };

            //foreach (var r in linq)
            //{
            //    Console.WriteLine(r.kmat + "\t " + r.naim);
            //}

            //Console.ReadLine();

            //return;

            //OleDbDataReader reader = command.ExecuteReader();
            //while (reader.Read())
            //    Console.WriteLine(String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}", 
            //                                    reader[0], 
            //                                    reader[1].ToString().Replace("?","і"), 
            //                                    reader[2], 
            //                                    reader[3], 
            //                                    Convert.ToDecimal(reader[4]), 
            //                                    Convert.ToInt32(reader[5]), 
            //                                    Convert.ToDecimal(reader[6])));

            //Console.OutputEncoding = Encoding.UTF8;

            RowTXT rowTXT = new RowTXT();
            rowTXT.CloneTable(dt);

            List<string> DoubleKmat = new List<string>();
            string n_kdk = "";

            try
            {
                int k = 0;
                foreach (DataRow row in dt.Rows)
                //    Console.WriteLine(String.Format("{0}", row.Field<string>("ceh").PadRight(11)));
                {

                    k++;
                    try
                    {
                        string ceh_s = row["ceh"].ToString();
                        string kmat_old = row["kmat"].ToString().Trim();
                        //string nnn = "001-";
                        //string kmat_old_format = kmat_old.Replace(" ", "").Replace("-", "").Replace(".", "");
                        //if (kmat_old_format == "")
                        //{
                        //    //throw new Exception("kmat_old_format = " + kmat_old_format);
                        //    ConvertKmat(kmat_old_format, ceh_s);
                        //}
                        string kmat = ConvertKmat(kmat_old, ceh_s, DoubleKmat);

                        bool flag1 = false;
                        object[] arrayColumn = row.ItemArray;

                        if (Flag(arrayColumn)) break;

                        int ceh = Convert.ToInt32(ceh_s);

                        n_kdk = row["n_kdk"].ToString();
                        string naim = row["naim"].ToString();
                        string size_type = row["size_type"].ToString();
                        int ei = FuncEI(row["ei"].ToString());
                        decimal price = FuncPrice(row["price"].ToString());
                        decimal count = FuncCount(row["count"].ToString());
                        decimal sum = FuncSum(row["sum"].ToString());

                        try
                        {
                            if (!DoubleKmat.Contains(kmat_old) | kmat_old == "")
                            {
                                DoubleKmat.Add(kmat_old);
                                //DataRow rowKmat = dt.AsEnumerable().Single(x => x["kmat"].ToString() == kmat_old);

                                //KSM.Add(rowKmat["kmat"].ToString(), rowKmat["kmat"].ToString());
                                //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}", k, ceh, kmat, kmat_old, ei, price, count, sum));

                                //int kk = 0;
                                //var rows = dt.Select("kmat = '" + kmat_old + "' ");

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
                                            //if (!KsmTable.IsRecord(kmat))
                                            //{
                                            //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);
                                            //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                            KSM.Add(kmat, ss2);
                                            Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                            //}
                                        }
                                        else
                                        {
                                            //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                            Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                        }

                                    }
                                }

                                //var rows2 = dt.Select("kmat ='" + kmat_old + "'");

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

                                                //if (!KsmTable.IsRecord(kmat))
                                                //{
                                                //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);
                                                //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                                KSM.Add(kmat, naim);
                                                Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                                //}
                                            }
                                            else
                                            {
                                                counter++;
                                                kmat = ConvertKmat("", ceh_s, DoubleKmat);

                                                //if (!KsmTable.IsRecord(kmat))
                                                //{
                                                //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);
                                                KSM.Add(kmat, naim);
                                                //}

                                                //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                                Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
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
                                        }

                                        KSM.Add(kmat, naim);
                                        //}

                                        //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                        Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                    }
                                }
                            }
                            else
                            {
                                //KSM.Add(kmat, naim);
                                Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                            }

                        }
                        catch (Exception ex)
                        {
                            //if (!KsmTable.IsRecord(kmat))
                            //{
                            //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                            //---------------------------------------------------------
                            if (kmat_old == "")
                            {
                                counter++;
                                kmat = ConvertKmat("", ceh_s, DoubleKmat);
                            }

                            KSM.Add(kmat, naim);
                            //}

                            //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                            Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}", k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));

                            //---------------------------------------------------------

                            //throw new Exception(ex.Message);
                            //DoubleKmat.Add(kmat_old);

                            //var rowsKmat = dt.AsEnumerable().Where(x => x["kmat"].ToString() == kmat_old);

                            //foreach (var rowKmat in rowsKmat)
                            //    rowTXT.Add(rowKmat, dt.Rows.IndexOf(rowKmat) + 2);

                        }

                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);

                        bool flag = false;
                        object[] arrayColumn = row.ItemArray;

                        for (int i = 0; i < arrayColumn.Length - 1; i++)
                            flag = arrayColumn[i].ToString() == "" ? true : false;

                        if (flag) break;

                        rowTXT.Add(row, k + 1);

                    }

                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);

            }

            Console.WriteLine("-------------------------------------------------");

            //foreach (var ss in KSM)
            //Console.WriteLine(ss.Key); 

            Console.WriteLine();

            decimal sumCount = dt.Select("n_kdk = '" + n_kdk.ToString() + "'").Sum(x => Convert.ToDecimal(x["count"]));
            //sumCountString = sumCount.ToString();

            Console.WriteLine("Сумма количество = " + sumCount);
            //rowTXT.WriteTXT(path, fileName, n_kdk_file);

            Console.ReadLine();

        }

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
            string ei_temp = ConvertEI(ei).ToString();

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

        public int ConvertEI(string ei)
        {
            switch (ei)
            {
                case "1":
                    return 839;
                case "2":
                    return 796;
                case "3":
                    return 166;
                case "5":
                    return 6;
                case "6":
                    return 761;
                case "7":
                    return 168;
            }

            return 796;
        }

        public void WriteTXT(string path, string fileName, string n_kdk)
        {

            //decimal positionCount = dt.AsEnumerable().Count(x => x["n_kdk"].ToString() == n_kdk);
            decimal positionCount = dt.Select("n_kdk = '" + n_kdk.ToString() + "'").Count();

            //decimal sum = dt.Select("n_kdk = '" + n_kdk.ToString() + "'").su;

            decimal sumCount = 0;
            string sumCountString = "";
            try
            {
                //sumCount = dt.AsEnumerable().Sum(x => Convert.ToDecimal(x["count"]));
                //sunCountString = sumCount.ToString();

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
                sb2.Append("К-во: " + CountPosition(positionCount) + new String('\t', 31) + "Всего: " + sumCountString);

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
}
