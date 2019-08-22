using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadC
{
	class Recource
	{
		static void GroupByListRecource()
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
	}
}
