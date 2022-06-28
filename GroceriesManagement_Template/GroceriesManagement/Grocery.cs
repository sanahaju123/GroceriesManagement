using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace GroceriesManagement
{
    public class Grocery : IComparable<Grocery>
    {
        public int GroceryId { get; set; }
        public String GroceryName { get; set; }
        public String Description { get; set; }
        public int price { get; set; }
        public DateTime ExpiryDate { get; set; }
        public Grocery(int groceryId, string groceryname, string description, int price, DateTime expiryDate)
        {
            this.GroceryId = groceryId;
            this.GroceryName = groceryname;
            this.Description = description;
            this.price = price;
            this.ExpiryDate = expiryDate;
        }



        public Grocery()
        {
        }

        public int CompareTo([AllowNull] Grocery other)
        {
            return this.GroceryName.CompareTo(other.GroceryName);
        }
    }
}

