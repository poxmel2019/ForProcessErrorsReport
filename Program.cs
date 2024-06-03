using Microsoft.Data.SqlClient;
using Spire.Xls;

namespace ForProcessErrors
{
    class Program
    {
        static async Task Main(string[] args)
        {

            Dictionary<int, string> months = new Dictionary<int, string>()
            {
                {1, "january"},
                {2, "february"},
                {3, "march"},
                {4, "april"},
                {5, "may"},
                {6, "june"},
                {7, "july"},
                {8, "august"},
                {9, "september"},
                {10, "october"},
                {11, "november"},
                {12, "december"}
            };

            int year = DateProcessing("y");
            Console.WriteLine(year);

            int month = DateProcessing("m");
            Console.WriteLine(month);

            //string month_name = new DateTime(year, month, 1).ToString("MMMM").ToLower(); // ru //?
            string month_name = months[month];

            Console.WriteLine(month_name);

            string address = "C:\\for_work\\db\\other\\three_important_reports\\errors_by_processes\\";
            string file = $"error_{month_name}_{year}";
            string dir = $"{address}{file}";

            if (!Directory.Exists(dir)) 
            {
                Directory.CreateDirectory(dir);
            }




            Workbook workbook = new Workbook();
            Worksheet worksheet_first = workbook.Worksheets[0];
            worksheet_first.Name = "Report";

            Worksheet worksheet_second = workbook.Worksheets[1];
            worksheet_second.Name = "ResultCodes";
                
            string connectionString = "Server=hbmssqltest.halykbank.nb;" +
                "Database=CorePayments;" +
                "User ID=CorePayments;" +
                "Password=0coayiwbYVReR;" +
                "TrustServerCertificate=true;";

            // querries
            // 1
            string all_processes = "select \r\ncount(*) all_processes\r\n" +
     "from processes p (nolock)\r\n" +
     "where StartDate between\r\n" +
     $"cast('{year}-{month-1}-01' as date) \r\n" +
     "and \r\n" +
     $"cast('{year}-{month}-01' as date)\r\n;";
            // 2
            string all_success_processes = "select \r\ncount(*)\r\n " +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                $"between\r\ncast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (16,37)\r\n;";    //81
            // 3
            string all_fail_process_exec = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                $"between\r\ncast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (10,35)\r\n;";
            // 4
            string all_success_exec_fail_callback = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (9, 34)\r\n;";
            // 5
            string all_fail_process_callback = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                $"between\r\ncast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (17,38)\r\n;";

            // payment
            // 6
            string payment_success_processes = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (16)\r\n;";
            // 7
            string payment_fail_check = "select \r\ncount(*) fail_check\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (13)\r\n;";
            // 8
            string payment_fail_exec = "select \r\ncount(*) fail_exec" +
                "\r\nfrom processes p (nolock)\r\n" +
                "where StartDate " +
                $"between\r\ncast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (10)\r\n;";
            // 9
            string payment_succes_exec_fail_callback = "select \r\n" +
                "count(*) no_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId = 9";
            // 10
            string payment_fail_callback = "select \r\ncount(*) " +
                "fail_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (17)";

            // tranfer
            // 11
            string transfer_success_process = "select \r\ncount(*) " +
                "success\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                $"between\r\ncast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (37)";
            // 12
            string transfer_fail_exec = "select \r\ncount(*) fail_exec" +
                "\r\nfrom processes p (nolock)\r\n" +
                "where StartDate " +
                $"between\r\ncast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (35)";
            // 13
            string transfer_success_exec = "select \r\ncount(*) " +         //81
                "no_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (34)\r\n;";
            // 14
            string transfer_fail_callback = "select \r\ncount(*) " +        
                "fail_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId = 38;";

            // 15
            string resultCodes = "select \r\nResultCode, " +                //81
                "count(resultcode) co\r\n" +
                "from ProcessOperations (nolock)\r\n" +
                "where StartDate between\r\n" +
                $"cast('{year}-{month-1}-01' as date) \r\n" +
                $"and \r\ncast('{year}-{month}-01' as date)\r\n" +
                "and ResultCode < 0\r\ngroup by ResultCode\r\n" +
                "order by ResultCode desc";


            string[] sqlExpressions =
            {
                all_processes, // 1
                "select '' as r",
                all_success_processes, // 2
                all_fail_process_exec, // 3
                all_success_exec_fail_callback, // 4
                all_fail_process_callback, // 5
                "select '' as r",
                payment_success_processes, // 6
                payment_fail_check, // 7
                payment_fail_exec, // 8
                payment_succes_exec_fail_callback, // 9
                payment_fail_callback, // 10
                "select '' as r",
                transfer_success_process, // 11
                transfer_fail_exec, // 12
                transfer_success_exec, // 13
                transfer_fail_callback // 14

            };

            string[] titles =
            {
                "Всего процессов", // 1
                "",
                "Успешных процессов", // 2
                "Неуспешных процессов на стадии выполнения", // 3
                "Неуспешных процессов на стадии успешного выполнения, " +
                "но без возврата коллбэка", // 4
                "Неуспешных процессов на стадии возврата коллбэка", // 5
                "",
                "Успешных платежей", // 6
                "Неуспешных платежей на стадии проверки", // 7
                "Неуспешных платежей на стадии выполнения", // 8
                "Неуспешных платежей на стадии успешного выполнения, " +
                "но без возврата коллбэка", // 9
                "Неуспешных платежей на стадии возврата коллбэка", // 10
                "",
                "Успешных переводов", // 11
                "Неуспешных переводов на стадии выполнения", // 12
                "Неуспешных переводов на стадии успешного выполнения, но без возврата коллбэка", // 13
                "Неуспешных переводов на стадии возврата коллбэка" // 14
            };
            
            using (SqlConnection connection = new SqlConnection(connectionString)) 
            {
                await connection.OpenAsync();

                for (int i = 0; i < sqlExpressions.Length; i++)
                {
                    SqlCommand command = new SqlCommand(sqlExpressions[i], connection);
                    SqlDataReader reader = await command.ExecuteReaderAsync();

                    if (reader.HasRows)
                    {
                        worksheet_first.Range[i + 1, 1].Value = titles[i];
                    }

                    int j = i+1;
                    while (await reader.ReadAsync())
                    {
                        object result = reader.GetValue(0);
                        if (reader.HasRows)
                        {
                            worksheet_first.Range[i + 1, 2].Value = result.ToString();
                        }
                        
                    }

                    await reader.CloseAsync();
                }

                Console.WriteLine("Query finished");

                // first string styling
                CellStyle style = workbook.Styles.Add("newStyle");
                style.Font.IsBold = true;
                for (int i = 1; i <= sqlExpressions.Length; i++)
                {
                    worksheet_first.Range[i, 1, i, 1].Style = style;
                }

                SqlCommand command2 = new SqlCommand(resultCodes, connection);
                SqlDataReader reader_second = await command2.ExecuteReaderAsync();

                worksheet_second.Range[1, 1].Value = "Код";
                worksheet_second.Range[1, 2].Value = "Количество";

                int k = 2;
                while (await reader_second.ReadAsync())
                {
                    object resultCode = reader_second.GetValue(0);
                    object quantity = reader_second.GetValue(1);

                    worksheet_second.Range[k, 1].Value = resultCode.ToString();
                    worksheet_second.Range[k, 2].Value = quantity.ToString();
                    k++;
                }

                await reader_second.CloseAsync();
                await connection.CloseAsync();

                Console.WriteLine("Result code finished");

                worksheet_second.Range[1, 1, 1, 2].Style = style;
                // fit width of columns
                worksheet_first.AllocatedRange.AutoFitColumns();

                Console.WriteLine("Excel file finished");

                // save to excel file
                try
                {
                    workbook.SaveToFile($"{dir}\\{file}.xlsx", ExcelVersion.Version2016);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                

            }

            Console.Read();

        }

        static int DateProcessing(string type)
        {
            string date_string, date_name;
            int date_int = 0;

            if (type == "y")
            {
                Console.WriteLine("Year:");
            }

            //else if (type == "m") 
            else
            {
                Console.WriteLine("Month:");
            }


            // year
            while (true)
            {
                date_string = Console.ReadLine();
                // check for void
                if (string.IsNullOrEmpty(date_string))
                {
                    if (type == "y")
                    {
                        date_int = DateTime.Now.Year;
                    }
                    else if (type == "m")
                    {
                        date_int = DateTime.Now.Month;
                    }

                    break;
                }
                else
                {   // check for number
                    bool isNumeric = int.TryParse(date_string, out date_int);
                    if (isNumeric)
                    {
                        date_int = Convert.ToInt32(date_string);
                        if (type == "y")
                        {
                            if (date_int < 0 || date_int > DateTime.Now.Year)
                            {
                                Console.WriteLine("Unreal!");
                            }
                            else
                            {
                                break;
                            }
                        }
                        else if (type == "m")
                        {
                            if (date_int < 1 || date_int > 12)
                            {
                                Console.WriteLine("Unreal!");
                            }
                            else
                            {
                                break;
                            }

                        }

                    }
                    else
                    {
                        Console.WriteLine("Number!");
                    }
                }
            }

            //Console.WriteLine(date_int);

            return date_int;
        }
    }
}