using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Documents;

namespace _4333Project
{
    public static class DataBaseInteractor
    {
        public static void Copy(string[,] list, SqlConnection sqlConnection)
        {
            
            string stringCommand = "INSERT INTO user (role, name, login, password) VALUES (@role, @name, @login, @password)";

            for (int i = 0; i < list.GetLength(0); i++)
            {
                int j = 0;

                using(SqlCommand command = new SqlCommand(stringCommand, sqlConnection))
                {
                    command.Parameters.AddWithValue("@role", 1);
                    command.Parameters.AddWithValue("@name", list[i, j + 1]);
                    command.Parameters.AddWithValue("@login", list[i, j + 2]);
                    command.Parameters.AddWithValue("@password", list[i, j + 3]);

                    command.ExecuteNonQuery();
                }
            }

        }
    }
}
