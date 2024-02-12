using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml.FormulaParsing;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace FirstProject
{
    public partial class Form1 : Form
    {
        private MySqlConnection databaseConnection;
        private DataTable dataTable;


        public Form1()
        {
            InitializeComponent();
        }

        private void ExportToExcel(DataGridView dataGridView1, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet 1");

                // Создаем словарь для хранения данных каждого пользователя
                Dictionary<string, List<object[]>> userData = new Dictionary<string, List<object[]>>();

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string userName = row.Cells["Ответственный"].Value != null ? row.Cells["Ответственный"].Value.ToString() : "";

                    // Если пользователя еще нет в словаре, добавляем его
                    if (!userData.ContainsKey(userName))
                    {
                        userData[userName] = new List<object[]>();
                    }

                    // Добавляем данные текущей строки для данного пользователя в список его данных
                    object[] rowData = new object[dataGridView1.Columns.Count];
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        rowData[j] = row.Cells[j].Value;
                    }
                    userData[userName].Add(rowData);
                }
                ExcelFont font = worksheet.Cells["A1"].Style.Font;
                font.Bold = true;
                font.Size = 20;
                font.Name = "Palatino Linotype";

                worksheet.Cells["A1:H1"].Merge = true;
                worksheet.Cells["A1:H1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1:H1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A1:H1"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                worksheet.Cells["A1:H1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1:H1"].Style.Fill.BackgroundColor.SetColor(Color.White);

                // Заполняем заголовки столбцов и номера столбцов
                int k = 1;
                for (int i = 1; i <= dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i].Value = "Расчет KPI для сотрудников ОПР департамента IT";
                    worksheet.Cells[2, i].Value = dataGridView1.Columns[i - 1].HeaderText;
                    worksheet.Cells[3, i].Value = k; k++;
                    worksheet.Cells[2, i].Style.Font.Name = "Palatino Linotype";
                    worksheet.Cells[3, i].Style.Font.Name = "Palatino Linotype";
                    worksheet.Cells[2, i].Style.Font.Size = 10;
                    worksheet.Cells[3, i].Style.Font.Size = 10;
                    worksheet.Cells[3, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[3, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[2, i].Style.Font.Bold = true;
                    worksheet.Cells[3, i].Style.Font.Bold = true;
                    worksheet.Cells[3, i].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    worksheet.Cells[2, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[2, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[2, i].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                }

                // Записываем данные для каждого пользователя в файл Excel
                int startRow = 3;
                foreach (var kvp in userData)
                {
                    // Записываем данные пользователя
                    foreach (var dataRow in kvp.Value)
                    {
                        startRow++;
                        for (int col = 0; col < dataRow.Length; col++)
                        {
                            worksheet.Cells[startRow, col + 1].Value = dataRow[col];
                            worksheet.Cells[startRow, col + 1].Style.Font.Size = 10;
                            worksheet.Cells[startRow, col + 1].Style.Font.Name = "Palatino Linotype";
                        }

                    }
                    worksheet.Cells[startRow, 1, startRow, 7].Value = kvp.Key;
                    worksheet.Cells[startRow, 1, startRow, 7].Merge = true;
                    worksheet.Cells[startRow, 1, startRow, 7].Style.Font.Size = 12;
                    worksheet.Cells[startRow, 1, startRow, 7].Style.Font.Bold = true;
                    worksheet.Cells[startRow, 1, startRow, 7].Style.Font.Name = "Palatino Linotype";
                    worksheet.Cells[startRow, 1, startRow, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    worksheet.Cells[startRow, 1, startRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, 1, startRow, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // Вычисляем и записываем сумму строк для данного пользователя в ячейку столбца
                    int sum = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        string userName = row.Cells["Ответственный"].Value?.ToString();
                        if (userName == kvp.Key)
                        {
                            sum++;
                        }
                    }
                    worksheet.Cells[startRow, 8].Value = sum;
                    worksheet.Cells[startRow, 8].Style.Font.Size = 12;
                    worksheet.Cells[startRow, 8].Style.Font.Bold = true;
                    worksheet.Cells[startRow, 8].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    worksheet.Cells[startRow, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, 8].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[startRow, 8].Style.Font.Name = "Palatino Linotype";

                    // Устанавливаем уровень группировки для строки
                    worksheet.Row(dataGridView1.ColumnCount).OutlineLevel = 1;

                    // Устанавливаем уровень группировки для строк этого пользователя
                    for (int i = startRow - kvp.Value.Count; i < startRow + 1; i++)
                    {
                        worksheet.Row(i).OutlineLevel = 1;
                    }

                    //Скрываем группированные строки
                    worksheet.Row(startRow - kvp.Value.Count).OutlineLevel = 0;
                    worksheet.Row(startRow - kvp.Value.Count).Collapsed = true;
                }

                // Определяем диапазон для форматирования
                int endColumn = 8;
                int endRow = startRow;

                // Форматируем диапазон ячеек как дату и время (если это не число)
                for (int i = 4; i <= endRow; i++)
                {
                    for (int j = 5; j <= endColumn; j++)
                    {
                        // Проверяем, что содержимое ячейки не является числом
                        if (!double.TryParse(worksheet.Cells[i, j].Value?.ToString(), out _))
                        {
                            worksheet.Cells[i, j].Style.Numberformat.Format = "dd/MM/yyyy HH:mm:ss";
                        }
                    }
                }

                int lastRow = worksheet.Dimension.End.Row;

                if (lastRow > 0)
                {
                    worksheet.DeleteRow(lastRow, 1); 
                    package.Save();
                }

                worksheet.Cells.AutoFitColumns();
                System.Threading.Thread.Sleep(1000);
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }
            MessageBox.Show("Данные успешно экспортированы в Excel!", "Экспорт данных", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void OkBtn_Click(object sender, EventArgs e)
        {
            string connectionString = ConfigurationManager.ConnectionStrings["MyConnectionString"].ConnectionString;

            databaseConnection = new MySqlConnection(connectionString);
            try
            {
                databaseConnection.Open();
                //string dateFrom = dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00";
                //string dateTo = dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59";

                string dateFrom = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss");
                string dateTo = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss");

                string query = $"SELECT number AS '№', service_items.name AS 'Сервис', bids.name AS 'Тема', contacts.name AS 'Ответственный'," +
                    $"created_at AS 'Дата регистрации', responded_date AS 'Время реакции', solution_provided_date AS 'Время разрешения'," +
                    $"closure_date AS 'Дата закрытия' FROM bids INNER JOIN service_items ON bids.service_item_id = service_items.id INNER JOIN contacts ON bids.owner_id = contacts.id WHERE created_at BETWEEN '{dateFrom}' AND '{dateTo}'";

                MySqlCommand commandDatabase = new MySqlCommand(query, databaseConnection);

                dataTable = new DataTable();

                using (MySqlDataAdapter adapter = new MySqlDataAdapter(commandDatabase))
                {
                    adapter.Fill(dataTable);
                }

                dataGridView1.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при получении данных: " + ex.Message);
            }
            finally
            {
                if (databaseConnection.State == ConnectionState.Open)
                    databaseConnection.Close();
            }
        }

        private void ExportToExelBtn_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    ExportToExcel(dataGridView1, saveFileDialog.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
