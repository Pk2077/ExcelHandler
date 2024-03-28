using FileHandler.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace FileHandler.Extensions
{
    public static class CustomerCrud
    {
        private static string _connectionString = 
            ConfigurationManager.ConnectionStrings["default"].ConnectionString;

        public static void InsertCustomers(DataRow row)
        {

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                    using (SqlCommand command = connection.CreateCommand())
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "InsertCustomers";

                        command.Parameters.AddWithValue("@CustomerName", row["Customer Name"]);
                        command.Parameters.AddWithValue("@CustomerCode", row["Customer Code"]);
                        command.Parameters.AddWithValue("@Address1", row["Add1"]);
                        command.Parameters.AddWithValue("@Address2", row["Add2"]);
                        command.Parameters.AddWithValue("@City", row["City"]);
                        command.Parameters.AddWithValue("@State", row["State Code"]);
                        command.Parameters.AddWithValue("@Pin", row["Pin"]);
                        command.Parameters.AddWithValue("@MobileNo", row["Mobile No"]);

                        command.ExecuteNonQuery();
                    }
            }
        }

        public static List<Customer> GetCustomers()
        {
            var Customers = new List<Customer>();
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "GetCustomers";

                    var reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Customers.Add(new Customer()
                        {
                            Id = reader.GetInt32(reader.GetOrdinal("Id")),
                            CustomerName = reader.GetString(reader.GetOrdinal("CustomerName")),
                            CustomerCode = reader.GetString(reader.GetOrdinal("CustomerCode")),
                            Address1 = reader.GetString(reader.GetOrdinal("Address1")),
                            Address2 = reader.GetString(reader.GetOrdinal("Address2")),
                            City = reader.GetString(reader.GetOrdinal("City")),
                            State = reader.GetString(reader.GetOrdinal("State")),
                            Pin = reader.GetString(reader.GetOrdinal("Pin")),
                            MobileNo = reader.GetString(reader.GetOrdinal("MobileNo")),
                        });
                    }
                    return Customers;
                }
            }
        }
        public static Customer GetCustomersByCode(string code)
        {
            var Customers = GetCustomers();
           return Customers.SingleOrDefault(c=>c.CustomerCode == code);
        }

        public static void DeleteCustomers(int Id)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    if (Id != 0)
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "DeleteCustomers";
                        command.Parameters.AddWithValue("@CustomerId", Id);
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                }
            }
        }

    }
}