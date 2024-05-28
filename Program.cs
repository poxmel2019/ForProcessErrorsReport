using Microsoft.Data.SqlClient;
using Spire.Xls;
using System.Runtime.ExceptionServices;

namespace ForProcessErrors
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            string connectionString = "Server=hbmssqltest.halykbank.nb;" +
                "Database=CorePayments;" +
                "User ID=CorePayments;" +
                "Password=0coayiwbYVReR;" +
                "TrustServerCertificate=true;";

            // all
            // 1
            string all_processes = "select \r\ncount(*) all_processes\r\n" +
     "from processes p (nolock)\r\n" +
     "where StartDate between\r\n" +
     "cast('2024-04-01' as date) \r\n" +
     "and \r\n" +
     "cast('2024-05-01' as date)\r\n;";
            // 2
            string all_success_processes = "select \r\ncount(*)\r\n " +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                "between\r\ncast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (16,37)\r\n;";    //81
            // 3
            string all_fail_process_exec = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\ncast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (10,35)\r\n;";
            // 4
            string all_success_exec_fail_callback = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (9, 34)\r\n;";
            // 5
            string all_fail_process_callback = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                "between\r\ncast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (17,38)\r\n;";

            // payment
            // 6
            string payment_success_processes = "select \r\ncount(*)\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (16)\r\n;";
            // 7
            string payment_fail_check = "select \r\ncount(*) fail_check\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (13)\r\n;";
            // 8
            string payment_fail_exec = "select \r\ncount(*) fail_exec" +
                "\r\nfrom processes p (nolock)\r\n" +
                "where StartDate between\r\ncast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (10)\r\n;";
            // 9
            string payment_succes_exec_fail_callback = "select \r\n" +
                "count(*) no_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId = 9";
            // 10
            string payment_fail_callback = "select \r\ncount(*) " +
                "fail_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (17)";

            // tranfer
            // 11
            string transfer_success_process = "select \r\ncount(*) " +
                "success\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate " +
                "between\r\ncast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\nand LastPortalServiceOperationStateId in (37)";
            // 12
            string transfer_fail_exec = "select \r\ncount(*) fail_exec" +
                "\r\nfrom processes p (nolock)\r\n" +
                "where StartDate " +
                "between\r\ncast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (35)";
            // 13
            string transfer_success_exec = "select \r\ncount(*) " +         //81
                "no_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId in (34)\r\n;";
            // 14
            string transfer_fail_callback = "select \r\ncount(*) " +        
                "fail_callback\r\n" +
                "from processes p (nolock)\r\n" +
                "where StartDate between\r\n" +
                "cast('2024-04-01' as date) \r\n" +
                "and \r\ncast('2024-05-01' as date)\r\n" +
                "and LastPortalServiceOperationStateId = 38";

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

                SqlCommand command = new SqlCommand(sqlExpressions[0], connection);
                SqlDataReader reader = await command.ExecuteReaderAsync();

                await reader.CloseAsync();


                for (int i = 0; i < sqlExpressions.Length; i++)
                {
                    command = new SqlCommand(sqlExpressions[i], connection);
                    reader = await command.ExecuteReaderAsync();

                    if (reader.HasRows)
                    {

                        worksheet.Range[i+1, 1].Value = titles[i];
                        
                    }

                    int j = i+1;
                    while (await reader.ReadAsync())
                    {
                        object result = reader.GetValue(0);
                        worksheet.Range[i+1, 2].Value = result.ToString();
                    }

                    await reader.CloseAsync();
                }

                Console.WriteLine("Finished");

                // first string styling
                CellStyle style = workbook.Styles.Add("newStyle");
                style.Font.IsBold = true;
                for (int i = 1; i <= sqlExpressions.Length; i++)
                {
                    worksheet.Range[i, 1, i, 1].Style = style;
                }

                // fit width of columns
                worksheet.AllocatedRange.AutoFitColumns();

                // save to excel file
                workbook.SaveToFile("C:\\for_work\\code\\my_projects\\ForExcel\\ReportError.xlsx",ExcelVersion.Version2016);

       
            }

            Console.Read();



        }
    }
}