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
        /// <param name="worksheet"></param>
        /// <param name="_schoolDetails"></param>
        /// <param name="_school"></param>
        /// <param name="i"></param>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool AddGroceryDetails(IXLWorksheet worksheet, Grocery _groceryDetails, List<Grocery> _grocery, int i, XLWorkbook workbook, string path)
        {
            bool res = false;
            try
            {
                _grocery.Add(_groceryDetails);

                worksheet.Cell(i, 1).Value = _grocery[0].GroceryId;
                worksheet.Cell(i, 2).Value = _grocery[0].GroceryName;
                worksheet.Cell(i, 3).Value = _grocery[0].Description;
                worksheet.Cell(i, 4).Value = _grocery[0].price;
                worksheet.Cell(i, 5).Value = _grocery[0].ExpiryDate;
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(path);
                    var content = stream.ToArray();
                }
                groceries.Add(_groceryDetails);
                res = true;
            }
            catch (Exception)
            {
                return res;
            }
            return res;
        }

        /// <summary>
        /// Serialize Groceries List object in Binary format and save it in text file
        /// </summary>
        /// <param name="pathTxt"></param>
        /// <param name="school"></param>
        /// <returns></returns>
        public static bool SerializeData(string pathTxt, List<Grocery> groceries)
        {
            bool res = false;
            try
            {
                var jsonValue = JsonConvert.SerializeObject(groceries);
                SaveTextFile(pathTxt, jsonValue);
                res = true;
            }
            catch (Exception)
            {
                return res;
            }
            return res;
        }

        /// <summary>
        /// Fetch all Groceries details from the saved text file .
        /// </summary>
        /// <param name="pathTxt"></param>
        /// <param name="jsonValue"></param>
        /// <returns></returns>
        public static bool SaveTextFile(string pathTxt, string jsonValue)
        {
            bool res = false;
            try
            {
                string text = File.ReadAllText(pathTxt);
                using (StreamWriter sw = File.AppendText(pathTxt))
                {
                    sw.WriteLine(jsonValue);
                }
                res = true;
            }
            catch (Exception)
            {
                return res;
            }
            return res;
        }

        /// <summary>
        /// Deserialize the fetched Grocery list object. 
        /// </summary>
        /// <param name="pathTxt"></param>
        public static bool DeserializeData(string pathTxt)
        {
            bool res = false;
            try
            {
                string txt = File.ReadAllText(pathTxt);
                var values = JsonConvert.DeserializeObject<List<Grocery>>(txt);
                DisplayAllDetails(values);
                res = true;
            }
            catch (Exception)
            {
                return res;
            }
            return res;
        }



        /// <summary>
        /// Show details of Groceries in descending order of name.
        /// </summary>
        /// <param name="values"></param>
        public static bool DisplayAllDetails(List<Grocery> values)
        {
            bool res = false;
            try
            {
                values.Reverse();
                foreach (Grocery skl in values)
                {
                    Console.WriteLine(skl.GroceryName);
                }
                res = true;
            }
            catch (Exception)
            {
                return res;
            }
            return res;
        }


        /// <summary>
        /// Empty text file.
        /// </summary>
        /// <param name="pathText"></param>
        /// <param name="fileMode"></param>
        public static bool ClearText(string pathText, FileMode fileMode)
        {
            bool res = false;
            try
            {

                using (var str = new FileStream(pathText, fileMode))
                {
                    using (var writer = new StreamWriter(str))
                    {
                        writer.Write("");
                    }
                }
                res = true;
            }
            catch (Exception ex)
            {
                return res;
            }
            return res;
        }
    }
}

