using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace GroceriesManagement.Tests.TestCases
{
    public class ExceptionalTests
    {
        // <summary>
        /// Creating referance variable of ITestOutputHelper and injecting in constructor
        /// </summary>
        private readonly ITestOutputHelper _output;
        private Grocery _grocery;
        private static string type = "Exception";
        public ExceptionalTests(ITestOutputHelper output)
        {
            _output = output;
            _grocery = new Grocery()
            {
                GroceryId = 1,
                GroceryName = "MyGrocery",
                Description = "Desc",
                price = 15,
                ExpiryDate = DateTime.Now
            };
        }
        /// <summary>
        /// If passing null will return false.
        /// </summary>
        /// <returns></returns>
        [Fact]
        public async Task<bool> TestFor_ReturnFalse_DisplayAllDetails()
        {
            //Arrange
            bool res = false;
            string testName; string status;
            List<Grocery> groceries = new List<Grocery>();
            groceries.Add(null);
            testName = CallAPI.GetCurrentMethodName();
            try
            {
                bool result = Program.DisplayAllDetails(groceries);
                //Act
                if (result == false)
                {
                    res = true;
                }
            }
            catch (Exception)
            {
                status = Convert.ToString(res);
                _output.WriteLine(testName + ":Failed");
                await CallAPI.saveTestResult(testName, status, type);
                return false;
            }
            ///Assert
            status = Convert.ToString(res);
            if (res == true)
            {
                _output.WriteLine(testName + ":Passed");
            }
            else
            {
                _output.WriteLine(testName + ":Failed");
            }
            await CallAPI.saveTestResult(testName, status, type);
            return res;
        }

    }
}

