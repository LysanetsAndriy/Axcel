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
        int CountColumn = 20;
        int CountRow = 50;
        FileManager fileManager = new FileManager();
        [JsonInclude]
        private Dictionary<string, cellProp> expressions = new Dictionary<string, cellProp>();
        public void SetExpression(string variableName, cellProp value)
        {
            expressions[variableName] = value;
        }
        private string LastUsedCell { get; set; }
        public MainPage()
        {
            InitializeComponent();
            CreateGrid();
            
        }
        private void CreateGrid()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }
        private void AddColumnsAndColumnLabels()
        {
            //Додати стовпці та підписи для стовпців
            for (int col = 0; col < CountColumn+1; col++)
            {
                grid.ColumnDefinitions.Add(new ColumnDefinition());
                if (col > 0)
                {
                    var label = new Label
                    {
                        Text = GetColumnName(col),
                        MinimumHeightRequest = 40,
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center
                    };
                    Grid.SetRow(label, 0); 
                    Grid.SetColumn(label, col);
                    grid.Children.Add(label);
                }
            }
        }
        private void AddRowsAndCellEntries()
        {
            
            grid.RowDefinitions.Add(new RowDefinition());
            //Додати рядки, підписи для рядків та комірки
            for (int row = 0; row < CountRow; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition());
                // Додати підписи для номера рядка
                var label = new Label
                {
                    Text = (row + 1).ToString(),
                    MinimumWidthRequest = 40,
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center
                };
                Grid.SetRow(label, row + 1);
                Grid.SetColumn(label, 0);
                grid.Children.Add(label);

                // Додати комірки (Entry) для вмісту
                for (int col = 0; col < CountColumn; col++)
                {
                    string stringCell = GetColumnName(col + 1) + (row + 1).ToString();

                    string cellText = "";
                    if (expressions.ContainsKey(stringCell) && expressions[stringCell].mode)
                    {
                        cellText = expressions[stringCell].formula;
                        var res = Calculator.Evaluate(cellText);
                        MyExcelMAUIAppVisitor.SetVariable(stringCell, (int)res);
                    }
                    else if (expressions.ContainsKey(stringCell))
                    {
                        cellText = expressions[stringCell].formula;
                        MyExcelMAUIAppVisitor.SetVariable(stringCell, (int)expressions[stringCell].value);
                    }
                    Entry cellEntry = new Entry
                    {
                        Text = cellText,
                        VerticalOptions = LayoutOptions.Center,
                        HorizontalOptions = LayoutOptions.Center,
                        MinimumWidthRequest = 260
                    };

                    cellEntry.Unfocused += Entry_Unfocused; // обробник події unfocused
                    Grid.SetRow(cellEntry, row + 1);
                    Grid.SetColumn(cellEntry, col + 1);
                    grid.Children.Add(cellEntry);
                }

            }
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

        // викликається, коли користувач вийде зі зміненої клітинки(втратить фокус)
        private async void Entry_Unfocused(object sender, FocusEventArgs e)
        {
            var entry = (Entry)sender;
            var row = Grid.GetRow(entry);
            var col = Grid.GetColumn(entry);
            string content = entry.Text;
            string stringCell = GetColumnName(col) + (row).ToString();

            if (Int32.TryParse(entry.Text, out int numValue1))
            {
                if (!expressions.ContainsKey(stringCell) ||
                    (MyExcelMAUIAppVisitor.IsNumberExpression(expressions[stringCell].formula) &&
                     expressions.ContainsKey(stringCell)))
                {
                    MyExcelMAUIAppVisitor.SetVariable(stringCell, numValue1);
                    entry.Focused += Entry_Focused;

                    SetExpression(stringCell, new cellProp(row, col, content, numValue1, false));
                    LastUsedCell = stringCell;
                }
                else if (!MyExcelMAUIAppVisitor.IsNumberExpression(expressions[stringCell].formula) &&
                         expressions.ContainsKey(stringCell) &&
                         (Calculator.Evaluate(expressions[stringCell].formula).ToString() != entry.Text))
                {
                    MyExcelMAUIAppVisitor.SetVariable(stringCell, numValue1);
                    entry.Focused += Entry_Focused;

                    SetExpression(stringCell, new cellProp(row, col, content, numValue1, false));
                    LastUsedCell = stringCell;
                }
            }
            else if (!String.IsNullOrWhiteSpace(entry.Text))
            {
                try
                {
                    var res = Calculator.Evaluate(entry.Text);
                    MyExcelMAUIAppVisitor.SetVariable(stringCell, (int)res);
                    entry.Text = res.ToString();
                    entry.Focused += Entry_Focused;

                    SetExpression(stringCell, new cellProp(row, col, content, (int)res, false));
                    LastUsedCell = stringCell;
                }
                catch (Exception ex)
                {
                    await DisplayAlert("Помилка", "Некоректний вираз!", "Ок");
                    entry.Text = entry.Text;
                }
            }
        }
        private void Entry_Focused(object sender, FocusEventArgs e)
        { 
            var entry = (Entry)sender;
            var row = Grid.GetRow(entry);
            var col = Grid.GetColumn(entry);
            var content = entry.Text;
            LastUsedCell = GetColumnName(col) + (row).ToString();
            entry.Focus();
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
                                DeleteColumnButton_Clicked(sender, e);
                            }
                            while (grid.RowDefinitions.Count > 1)
                            {
                                DeleteRowButton_Clicked(sender, e);
                            }
                            LastUsedCell = "";
                            expressions = tableInfo.Expressions;
                            CountColumn = tableInfo.CountColumn - 1;
                            CountRow = tableInfo.CountRow - 1;
                            CreateGrid();
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
            if (!string.IsNullOrEmpty(LastUsedCell) && expressions.ContainsKey(LastUsedCell))
            {
                if (!expressions[LastUsedCell].mode) // current mode - value
                {
                    expressions[LastUsedCell].mode = !expressions[LastUsedCell].mode;
                    int row = expressions[LastUsedCell].row;
                    int col = expressions[LastUsedCell].col;
                    foreach (var child in grid.Children)
                    {
                        if (grid.GetRow(child) == row && grid.GetColumn(child) == col)
                        {
                            if (child is Entry entry)
                            {
                                // Found the entry in the specified row and column
                                entry.Text = expressions[LastUsedCell].formula;
                                break;
                            }
                        }
                    }
                }
                else // current mode - formula
                {
                    expressions[LastUsedCell].mode = !expressions[LastUsedCell].mode;
                    int row = expressions[LastUsedCell].row;
                    int col = expressions[LastUsedCell].col;
                    foreach (var child in grid.Children)
                    {
                        if (grid.GetRow(child) == row && grid.GetColumn(child) == col)
                        {
                            if (child is Entry entry)
                            {
                                // Found the entry in the specified row and column
                                entry.Text = expressions[LastUsedCell].value.ToString();
                                break;
                            }
                        }
                    }
                }
            }
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
        private int GetRowIndexFromCellIndex(string cellIndex)
        {
            // Extract and convert the row index from the cell index string
            string rowIndexString = Regex.Replace(cellIndex, "[^0-9]", "");
            int rowIndex = int.Parse(rowIndexString) - 1; // Subtract 1 because row index is 0-based
            return rowIndex;
        }
        private void DeleteRowButton_Clicked(object sender, EventArgs e)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;

                List<string> keysToRemove = new List<string>();
                foreach (var entry in expressions)
                {
                    string cellIndex = entry.Key;
                    int rowIndex = GetRowIndexFromCellIndex(cellIndex);
                    if (rowIndex >= lastRowIndex-1)
                    {
                        keysToRemove.Add(cellIndex);
                    }
                }

                foreach (string key in keysToRemove)
                {
                    expressions.Remove(key);
                }

                grid.RowDefinitions.RemoveAt(lastRowIndex);
                var children = grid.Children.ToList();
                foreach (var child in children.Where(child => grid.GetRow(child) == lastRowIndex))
                {
                    grid.Children.Remove(child);
                }
                
            }
        }
        private int GetColumnIndexFromCellIndex(string cellIndex)
        {
            // Extract and convert the column index from the cell index string
            string columnIndexString = Regex.Replace(cellIndex, "[^A-Za-z]", "");
            int columnIndex = columnIndexString[0] - 'A'; // Convert to a 0-based index
            return columnIndex;
        }
        private void DeleteColumnButton_Clicked(object sender, EventArgs e)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;

                List<string> keysToRemove = new List<string>();
                foreach (var entry in expressions)
                {
                    string cellIndex = entry.Key;
                    int columnIndex = GetColumnIndexFromCellIndex(cellIndex);
                    if (columnIndex >= lastColumnIndex-1)
                    {
                        keysToRemove.Add(cellIndex);
                    }
                }

                foreach (string key in keysToRemove)
                {
                    expressions.Remove(key);
                }

                grid.ColumnDefinitions.RemoveAt(lastColumnIndex);

                var children = grid.Children.ToList();
                foreach (var child in children.Where(child => grid.GetColumn(child) == lastColumnIndex))
                {
                    grid.Children.Remove(child);
                }
            }
        }
        private void AddRowButton_Clicked(Object sender, EventArgs e)
        {
            int newRow = grid.RowDefinitions.Count;

            //Add a new row definition
            grid.RowDefinitions.Add(new RowDefinition());

            //Add label for the row number
            var label = new Label
            {
                Text = newRow.ToString(),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center,
                MinimumWidthRequest = 40
            };
            Grid.SetRow(label, newRow);
            Grid.SetColumn(label, 0);
            grid.Children.Add(label);
            
            //Add entry cells for the new row
            for (int col = 0; col < grid.ColumnDefinitions.Count; col++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    MinimumWidthRequest = 260
                };
                Grid.SetRow(entry, newRow);
                Grid.SetColumn(entry, col + 1);
                grid.Children.Add(entry);
            }
        }
        private void AddColumnButton_Clicked(object sender, EventArgs e)
        {
            int newColumn = grid.ColumnDefinitions.Count;
            // Add a new column definition
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            // Add label for the column name
            var label = new Label
            {
                Text = GetColumnName(newColumn),
                VerticalOptions = LayoutOptions.Center,
                HorizontalOptions = LayoutOptions.Center,
                MinimumHeightRequest = 40
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, newColumn);
            grid.Children.Add(label);
            // Add entry cells for the new column
            for (int row = 0; row < grid.RowDefinitions.Count; row++)
            {
                var entry = new Entry
                {
                    Text = "",
                    VerticalOptions = LayoutOptions.Center,
                    HorizontalOptions = LayoutOptions.Center,
                    MinimumWidthRequest = 260

                };
                entry.Unfocused += Entry_Unfocused;
                Grid.SetRow(entry, row + 1);
                Grid.SetColumn(entry, newColumn);
                grid.Children.Add(entry);
            }
        }
    }
}