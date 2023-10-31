using System.Data.Common;
using System.Runtime.InteropServices;
using Microsoft.Maui.Storage;
using CommunityToolkit.Maui.Storage;
using System.Linq.Expressions;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text;
using CommunityToolkit.Maui.Converters;
using System.Text.RegularExpressions;

namespace MyExcelMAUIApp
{
    public class cellProp
    {
        [JsonInclude]
        public int row { get; set; }
        [JsonInclude]
        public int col { get; set; }
        [JsonInclude]
        public string formula { get; set; }
        [JsonInclude]
        public int value { get; set; }
        [JsonInclude]
        public bool mode { get; set; } // false - value, true - formula

        public cellProp(int row, int col, string formula, int value, bool mode)
        {
            this.row = row;
            this.col = col;
            this.formula = formula;
            this.value = value;
            this.mode = mode;
        }
    }
    public partial class MainPage : ContentPage
    {
        public int CountColumn = 20;
        public int CountRow = 50; 
        FileManager fileManager = new FileManager();
        [JsonInclude]
        private Dictionary<string, cellProp> expressions = new Dictionary<string, cellProp>();
        public void SetExpression(string variableName, cellProp value)
        {
            expressions[variableName] = value;
        }
        private Table table;
        public MainPage()
        {
            InitializeComponent();
            table = new Table(grid, expressions); // Initialize the Table instance
            table.Initialize();
            
        }
        private string GetColumnName(int colIndex)
        {
            int dividend = colIndex;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(modulo + 65) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
        private async void SaveAsButton_Clicked(object sender, EventArgs e)
        {
            try
            {
                string result = await DisplayPromptAsync("Enter file name", "File name: ");
                if (string.IsNullOrEmpty(result)) // Check if the user entered a file name
                {
                    throw new ArgumentNullException();
                }

                var folder = await FolderPicker.PickAsync(default);

                string filePath = Path.Combine(folder.Folder.Path, result + ".json");
                fileManager.SaveAs(expressions, grid.ColumnDefinitions.Count, grid.RowDefinitions.Count, filePath);
                await DisplayAlert("Збережено", "Таблиця збережена успішно!", "Ок");
            }
            catch (Exception ex)
            {
                using (StreamWriter writer = new StreamWriter("D:\\rlv\\console.txt"))
                {
                    writer.WriteLine($"Exception: {ex.Message}");
                    writer.WriteLine($"Stack Trace: {ex.StackTrace}");
                    if (ex is ArgumentNullException)
                        await DisplayAlert("Помилка", "Ім'я файлу не може бути порожнім!", "Ок");
                    else
                        await DisplayAlert("Помилка", "Помилка під час збереження файлу!", "Ок");
                }
            }
        }
        private async void SaveButton_Clicked(object sender, EventArgs e)
        {
            try
            {
                fileManager.Save(expressions, grid.ColumnDefinitions.Count, grid.RowDefinitions.Count);
            }
            catch (Exception ex)
            {
                await DisplayAlert("Помилка", "Помилка під час збереження файлу!", "Ок");
            }
        }
        private async void LoadButton_Clicked(object sender, EventArgs e)
        {
                var loadPicker = new PickOptions()
                {
                    FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.WinUI, new[] { ".json" } },
                }),
                };

                var file = await FilePicker.PickAsync(loadPicker);

                if (file != null)
                {
                    try
                    {
                        string filePath = file.FullPath;
                        TableInfo tableInfo = fileManager.Load(filePath);

                        if (tableInfo != null)
                        {
                            while (grid.ColumnDefinitions.Count > 1)
                            {
                                table.DeleteColumn(sender);
                            }
                            while (grid.RowDefinitions.Count > 1)
                            {
                                table.DeleteRow(sender);
                            }
                            table.LastUsedCell = "";
                            expressions = tableInfo.Expressions;
                            CountColumn = tableInfo.CountColumn - 1;
                            CountRow = tableInfo.CountRow - 1;
                            table.Initialize();
                            await DisplayAlert("Завантажено", "Таблиця завантажена успішно!", "Ок");
                        }
                    }
                    catch (Exception ex)
                    {
                        await DisplayAlert("Помилка", "Помилка під час завантаження таблиці!", "Ок");
                    }
                }
        }
        private void ModeButton_Clicked(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(table.LastUsedCell) && expressions.ContainsKey(table.LastUsedCell))
            {
                if (!expressions[table.LastUsedCell].mode) // current mode - value
                {
                    expressions[table.LastUsedCell].mode = !expressions[table.LastUsedCell].mode;
                    int row = expressions[table.LastUsedCell].row;
                    int col = expressions[table.LastUsedCell].col;
                    foreach (var child in grid.Children)
                    {
                        if (grid.GetRow(child) == row && grid.GetColumn(child) == col)
                        {
                            if (child is Entry entry)
                            {
                                // Found the entry in the specified row and column
                                entry.Text = expressions[table.LastUsedCell].formula;
                                break;
                            }
                        }
                    }
                }
                else // current mode - formula
                {
                    expressions[table.LastUsedCell].mode = !expressions[table.LastUsedCell].mode;
                    int row = expressions[table.LastUsedCell].row;
                    int col = expressions[table.LastUsedCell].col;
                    foreach (var child in grid.Children)
                    {
                        if (grid.GetRow(child) == row && grid.GetColumn(child) == col)
                        {
                            if (child is Entry entry)
                            {
                                // Found the entry in the specified row and column
                                entry.Text = expressions[table.LastUsedCell].value.ToString();
                                break;
                            }
                        }
                    }
                }
            }
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            table.DeleteRow(sender);
        }
        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        { 
            table.DeleteColumn(sender);
        }
        private void AddRowButton_Clicked(Object sender, EventArgs e)
        { 
            table.AddRow(sender);
        }
        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            table.AddColumn(sender);
        }
        private async void ExitButton_Clicked(object sender, EventArgs e)
        {
            bool answer = await DisplayAlert("Підтвердження", "Ви дійсно хочете вийти?", "Так", "Ні");
            if (answer)
            {
               System.Environment.Exit(0);
            }
        }
        private async void HelpButton_Clicked(object sender, EventArgs e)
        {
            await DisplayAlert("Довідка", "Лабораторна робота 1. Студента групи К-24 Андрія Лисанця\n" +
                "Особливості виразів для варіанту:\n" +
                "Вираз обов'язково має містити один із цих знаків: =, <, >;\n" +
                "Резульатом виразу є число 1 або 0;\n" +
                "Клітинка може містити лише число або вираз", "Ок");
        }
    }
}