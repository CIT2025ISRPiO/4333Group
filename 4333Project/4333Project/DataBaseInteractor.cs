using System;
using System.Data.SqlClient;
using System.Windows.Documents;

namespace _4333Project
{
    public static class DataBaseInteractor
    {
        public static void Add(string tablename)
        {

                using(SqlCommand command = new SqlCommand($"INSERT INTO {tableName}", connectionString)) {

                    command.ExecuteNonQuery

                    for(int i = 0; i < _rows; i++)
                    {
                        usersEntities.Users.Add(new Users()
                        {
                            Log =
                        list[i, 1],
                            Pass = list[i, 2]
                        });
                    }
                usersEntities.SaveChanges();
            }

        }
    }
}
