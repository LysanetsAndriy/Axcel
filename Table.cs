using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MyExcelMAUIApp
{
    public class Table : MainPage
    {
        private Grid grid;
        private Dictionary<string, cellProp> expressions;
        public string LastUsedCell { get; set; }

        public Table(Grid grid, Dictionary<string, cellProp> expressions)
        {
            this.grid = grid;
            this.expressions = expressions;
        }

        public void Initialize()
        {
            AddColumnsAndColumnLabels();
            AddRowsAndCellEntries();
        }
        public void AddColumnsAndColumnLabels()
        {
            for (int col = 0; col < CountColumn + 1; col++)
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
        public void Entry_Focused(object sender, FocusEventArgs e)
        {
                var entry = (Entry)sender;
                var row = Grid.GetRow(entry);
                var col = Grid.GetColumn(entry);
                var content = entry.Text;
                LastUsedCell = GetColumnName(col) + (row).ToString();
                entry.Focus();
        }
        private int GetRowIndexFromCellIndex(string cellIndex)
        {
            // Extract and convert the row index from the cell index string
            string rowIndexString = Regex.Replace(cellIndex, "[^0-9]", "");
            int rowIndex = int.Parse(rowIndexString) - 1; // Subtract 1 because row index is 0-based
            return rowIndex;
        }
        public void DeleteRow(object sender)
        {
            if (grid.RowDefinitions.Count > 1)
            {
                int lastRowIndex = grid.RowDefinitions.Count - 1;

                List<string> keysToRemove = new List<string>();
                foreach (var entry in expressions)
                {
                    string cellIndex = entry.Key;
                    int rowIndex = GetRowIndexFromCellIndex(cellIndex);
                    if (rowIndex >= lastRowIndex - 1)
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
        public int GetColumnIndexFromCellIndex(string cellIndex)
        {
            // Extract and convert the column index from the cell index string
            string columnIndexString = Regex.Replace(cellIndex, "[^A-Za-z]", "");
            int columnIndex = columnIndexString[0] - 'A'; // Convert to a 0-based index
            return columnIndex;
        }
        public void DeleteColumn(object sender)
        {
            if (grid.ColumnDefinitions.Count > 1)
            {
                int lastColumnIndex = grid.ColumnDefinitions.Count - 1;

                List<string> keysToRemove = new List<string>();
                foreach (var entry in expressions)
                {
                    string cellIndex = entry.Key;
                    int columnIndex = GetColumnIndexFromCellIndex(cellIndex);
                    if (columnIndex >= lastColumnIndex - 1)
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
        public void AddRow(Object sender)
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
        public void AddColumn(object sender)
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
