using System;
using System.Data;
using System.Data.SqlClient;

namespace ImportRowCS
{
	class Class1
	{
		static void Main11(string[] args)
		{
			// 
			// TODO: Add code to start application here.
			DataTable tblProducts = new DataTable();
			DataTable tblProductsCopy = new DataTable();

			int tblProductsCount;
			int tblProductsCopyCount;
			int i;
			
			SqlConnection Conn = new SqlConnection("Server=(local);database=Northwind;UID=<User ID>;PWD=<Password>");
			
			SqlDataAdapter da = new SqlDataAdapter("Select * from products", Conn);
			
			DataSet ds = new DataSet();
			da.Fill(ds, "products");
			tblProducts = ds.Tables["products"];
			tblProductsCount = tblProducts.Rows.Count;

			
			Console.WriteLine("Table tblProducts has " + tblProductsCount.ToString() + " Rows");

			for (i = 0; i <= 4; ++i)
			{
				Console.WriteLine("Row(" + i.ToString() + ") = " + tblProducts.Rows[i][1]);
			}

			tblProductsCopy = tblProducts.Clone();

			// Use the ImportRow method to copy from Products table to its clone.
			for (i = 0; i <= 4; ++i)
			{
				tblProductsCopy.ImportRow(tblProducts.Rows[i]);
			}
			tblProductsCopyCount = tblProductsCopy.Rows.Count;
			// Write blank line.
			Console.WriteLine();
			// Write the number of rows in tblProductsCopy table to the screen.
			Console.WriteLine("Table tblProductsCopy has " +
							  tblProductsCopyCount.ToString() + " Rows");

			// Loop through the top five rows, and write the first column to the screen.
			for (i = 0; i <= tblProductsCopyCount - 1; ++i)
			{
				Console.WriteLine("Row(" + i.ToString() + ") = " + tblProductsCopy.Rows[i][1]);
			}

			// This line keeps the console open until you press ENTER.
			Console.ReadLine();

		}
	}
}