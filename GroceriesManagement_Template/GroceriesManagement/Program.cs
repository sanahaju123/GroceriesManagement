using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace GroceriesManagement
{
    public class Program
    {
        private static List<Grocery> groceries = new List<Grocery>();
        public static void Main()
        {
            FileInfo fileInfo = new FileInfo(AppDomain.CurrentDomain.BaseDirectory);
            string parentDir = fileInfo.Directory.Parent.Parent.Parent.Parent.ToString();
            string path = Path.Combine(parentDir, @"GroceriesManagement/Assets/ResultSheet.xlsx");
            string pathTxt = Path.Combine(parentDir, @"GroceriesManagement/Assets/Test.txt");


            ClearText(pathTxt, FileMode.Truncate);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Grocery Details");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "GroceryId";
                worksheet.Cell(currentRow, 2).Value = "GroceryName";
                worksheet.Cell(currentRow, 3).Value = "Description";
                worksheet.Cell(currentRow, 4).Value = "price";
                worksheet.Cell(currentRow, 5).Value = "ExpiryDate";
                for (int i = 2; i < 7; i++)
                {
                    Console.Write("GroceryId: ");
                    int ID = int.Parse(Console.ReadLine());

                    Console.Write("GroceryName: ");
                    string name = Console.ReadLine();

                    Console.Write("Description: ");
                    string address = Console.ReadLine();

                    Console.Write("price: ");
                    int price = int.Parse(Console.ReadLine());

                    Console.Write("ExpiryDate: ");
                    DateTime date = DateTime.Parse(Console.ReadLine());

                    Grocery _groceryDetails = new Grocery(ID, name, address, price, date);
                    List<Grocery> _grocery = new List<Grocery>();

                    AddGroceryDetails(worksheet, _groceryDetails, _grocery, i, workbook, path);

                }
                SerializeData(pathTxt, groceries);
                DeserializeData(pathTxt);

            }
        }

        /// <summary>
        /// Add at least 5 Groceries details using List generic collection. <userinput>
        /// </summary>
        
        public static bool AddGroceryDetails(IXLWorksheet worksheet, Grocery _groceryDetails, List<Grocery> _grocery, int i, XLWorkbook workbook, string path)
        {
            //Write Your Code Here
            throw new NotImplementedException();
        }

        /// <summary>
        /// Serialize Groceries List object in Binary format and save it in text file
        /// </summary>
        
        public static bool SerializeData(string pathTxt, List<Grocery> groceries)
        {
            //Write Your Code Here
            throw new NotImplementedException();
        }

        /// <summary>
        /// Fetch all Groceries details from the saved text file .
        /// </summary>
       
        public static bool SaveTextFile(string pathTxt, string jsonValue)
        {
            //Write Your Code Here
            throw new NotImplementedException();
        }

        /// <summary>
        /// Deserialize the fetched Grocery list object. 
        /// </summary>
        public static bool DeserializeData(string pathTxt)
        {
            //Write Your Code Here
            throw new NotImplementedException();
        }



        /// <summary>
        /// Show details of Groceries in descending order of name.
        /// </summary>
        public static bool DisplayAllDetails(List<Grocery> values)
        {
            //Write Your Code Here
            throw new NotImplementedException();
        }


        /// <summary>
        /// Empty text file.
        /// </summary>
        public static bool ClearText(string pathText, FileMode fileMode)
        {
            //Write Your Code Here
            throw new NotImplementedException();
        }
    }
}

