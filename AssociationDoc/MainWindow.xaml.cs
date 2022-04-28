using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;


namespace AssociationDoc
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ListViewSelectedFiles.ItemsSource = items;
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Выберите документы для объединения (можно несколько)";
            openFileDialog.Filter = "Таблицы (*.xlsx,*.csv,*.xls)|*.xlsx;*.csv;*.xls"; //форматы файлов, которые отображаются при выборе
            if (openFileDialog.ShowDialog() == true)
            {
                AddFiles(openFileDialog.FileNames);
                string[] paths = openFileDialog.FileNames;
            }
        }

        List<FileSource> items = new List<FileSource>();

        private void ListViewSelectedFiles_PreviewDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                AddFiles(files);
            }
        }

        public void AddFiles(string[] files)
        {
            try
            {

                string except = "";
                foreach (string file in files)
                {
                    string docType = file.Substring(file.Length - 5, 5);
                    Regex regexDocType = new Regex(@"(\.csv)|(\.xlsx)|(\.xls)");
                    if (regexDocType.IsMatch(docType))
                    {
                        if (items.Where(x => x.Path == file).Count() < 1)
                        {
                            items.Add(new FileSource() { Path = file, FileName = Path.GetFileName(file) });
                            ListViewSelectedFiles.Items.Refresh();
                        }
                    }
                    else
                    {
                        except = "Ошибка, можно добавлять только файлы .xlsx, .xls, .csv.\nОстальные файлы не были добавлены";

                    }

                }
                if (except.Length > 0)
                {
                    throw new Exception(except);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public bool Unlock(Excel.Worksheet sheet, string fileName) //проверка на защиту листа и ввод пароля
        {
            if (sheet.ProtectContents)
            {
                if (!Password.allPasswords) //если защищен
                {
                    while (true)
                    {
                        GetPasswordWindow g = new GetPasswordWindow(fileName); //окно для ввода пароля
                        g.Owner = this;
                        g.WindowStartupLocation = WindowStartupLocation.CenterOwner; //центровка
                        g.ShowDialog();
                        if (unlockSheet(sheet)) //если разблокрировался, то спрашиваем применять ли пароль ко всем
                        {
                            MessageBoxResult r = MessageBox.Show("Этот пароль подходит ко всем следующим файлам?", "Вопрос", MessageBoxButton.YesNo, MessageBoxImage.Question);
                            switch (r)
                            {
                                case MessageBoxResult.Yes: Password.allPasswords = true; break;
                                default: Password.allPasswords = false; break;
                            }
                            return true;
                        }
                        else //иначе предлагаем ввести заново или же отменить объединение
                        {
                            Password.allPasswords = false;
                            MessageBoxResult r = MessageBox.Show("Пароль не подошел!\nДа - ввести ещё раз\nНет - отменить объединение", "Вопрос", MessageBoxButton.YesNo, MessageBoxImage.Question);
                            switch (r)
                            {
                                case MessageBoxResult.Yes: break;
                                default: return false;
                            }
                        }
                    }
                }
                else
                {
                    if (unlockSheet(sheet))
                    {
                        return true;
                    }
                    else
                    {
                        Password.allPasswords = false;
                        return Unlock(sheet, fileName);
                    }
                }
            }
            else
            {
                return true;
            }

        }

        public bool unlockSheet(Excel.Worksheet sheet) //вводим пароль для разблокировки листа
        {
            try
            {
                sheet.Unprotect(Password.password);
                return true;
            }
            catch
            {
                Password.allPasswords = false;
                return false;
            }
        }

        private void Association_Click(object sender, RoutedEventArgs e)
        {
            if (endFile.Path != null && items.Count > 0)
            {
                string FilesNotExists = "Следующие файлы были перемещены или удалены:\n";
                bool filesExists = true;

                if (!File.Exists(endFile.Path))
                {
                    filesExists = false;
                    FilesNotExists += endFile.Path + "\n";
                }
                foreach (FileSource file in items)
                {
                    if (!File.Exists(file.Path))
                    {
                        filesExists = false;
                        FilesNotExists += file.Path + "\n";
                    }
                }
                if (!filesExists)
                {
                    MessageBox.Show(FilesNotExists, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                Excel.Application xlApp = new Excel.Application(); //Excel
                Excel.Workbook wbStartExcel = xlApp.Workbooks.Open(items[0].Path); //название файла Excel откуда будем копировать лист 
                Excel.Worksheet wsStartExcel = wbStartExcel.Worksheets[1]; //лист Excel, откуда будем копировать данные
                Excel.Workbook wbNewExcel = xlApp.Workbooks.Open(endFile.Path);  //рабочая книга, в которую будем вставлять данные            
                Excel.Worksheet wsNewExcel = wbNewExcel.Sheets[1]; //первый лист по порядку - в него будем вставлять данные; //лист Excel на который будем копировать
                Excel.Workbook wbTitle = xlApp.Workbooks.Open(Environment.CurrentDirectory + "\\Title.xlsx"); //шапка документа
                Excel.Worksheet wsTitle = wbTitle.Sheets[1];
                if (Unlock(wsStartExcel, items[0].FileName))
                {
                    int idLastCopy = wsStartExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    wbStartExcel.Close(false);
                    try
                    {
                        foreach (FileSource file in items)
                        {
                            wbStartExcel = xlApp.Workbooks.Open(file.Path);
                            wsStartExcel = wbStartExcel.Worksheets[1]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
                            if (!Unlock(wsStartExcel, file.FileName))
                            {
                                throw new Exception("Вы отменили объединение файлов!");
                            }

                            idLastCopy = wsStartExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row; //ид последней записи
                            if (wsNewExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row == 1) //если первая запись в новом листе, то 
                            {                                                                               //копируем шапку
                                wsTitle.Range["A1:CU26"].Copy();
                                wsNewExcel.Range["A1"].PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);
                                wsNewExcel.Range["A1"].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                                ColumnWidths(wsNewExcel);
                                wbNewExcel.Save();
                            }
                            if (idLastCopy > 26)
                            {
                                wsStartExcel.Range["A27:CU" + idLastCopy].Copy();
                                int id = wsNewExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 2; //оставили место под название
                                wsNewExcel.Range["A" + id].PasteSpecial(Excel.XlPasteType.xlPasteAll);

                                id--;
                                wsNewExcel.Range["A" + id].Value = wsStartExcel.Range["V7"].Text; //запись названия
                                wsNewExcel.get_Range("A" + id, "CU" + id).Merge(Type.Missing);
                                wsNewExcel.Range["A" + id].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                wsNewExcel.Range["A" + id].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                wsNewExcel.Range["A" + id].Font.Size = 30;
                            }
                            wbStartExcel.Close(false);
                        }
                        int idLastNew = wsNewExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                        wsNewExcel.Range["AK24"].Value = "=СУММ(AK25:AK" + idLastNew + ")"; //переделываем формулы
                        wsNewExcel.Range["AR24"].Value = "=СУММ(AR25:AR" + idLastNew + ")";
                        wsNewExcel.Range["AY24"].Value = "=СУММ(AY25:AY" + idLastNew + ")";
                        wsNewExcel.Range["BF24"].Value = "=СУММ(BF25:BF" + idLastNew + ")";
                        wsNewExcel.Range["BM24"].Value = "=СУММ(BM25:BM" + idLastNew + ")";
                        wsNewExcel.Range["BT24"].Value = "=СУММ(BT25:BT" + idLastNew + ")";
                        wsNewExcel.Range["CA24"].Value = "=СУММ(CA25:CA" + idLastNew + ")";
                        wsNewExcel.Range["CH24"].Value = "=СУММ(CH25:CH" + idLastNew + ")";
                        wsNewExcel.Range["CO24"].Value = "=СУММ(CO25:CO" + idLastNew + ")";

                        if (idLastCopy > 26)
                        {
                            for (int ii = 27; ii <= idLastNew; ii++)
                            {
                                wsNewExcel.Range["A" + ii].RowHeight = 60;
                            }
                        }
                        wbNewExcel.Close(true); //закрываем всё
                        wbTitle.Close(false);
                        xlApp.Quit();
                        GC.Collect();

                        Password.allPasswords = false;

                        Process.Start(endFile.Path);
                    }
                    catch (Exception ex) //пробуем всё закрыть в случае ошибки
                    {
                        try
                        {
                            wbStartExcel.Close(false);
                        }
                        catch { }
                        try
                        {
                            wbNewExcel.Close(false);
                        }
                        catch { }
                        try
                        {
                            wbTitle.Close(false);
                        }
                        catch { }
                        try
                        {
                            xlApp.Quit();
                            GC.Collect();
                        }
                        catch { }
                        Password.allPasswords = false;

                        MessageBox.Show("Произошла ошибка! Пропробуйте ещё раз.\nКод ошибки: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

                    }
                }
                else
                {
                    try
                    {
                        wbStartExcel.Close(false);
                    }
                    catch { }
                    try
                    {
                        wbNewExcel.Close(false);
                    }
                    catch { }
                    try
                    {
                        wbTitle.Close(false);
                    }
                    catch { }
                    try
                    {
                        xlApp.Quit();
                        GC.Collect();
                    }
                    catch { }
                    Password.allPasswords = false;
                    MessageBox.Show("Вы отменили объединение файлов!", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

            }
            else
            {

                MessageBox.Show("Не хватает файлов для объединения", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

        public void ColumnWidths(Excel.Worksheet wsNewExcel) //ширина столбцов
        {
            wsNewExcel.Columns.ColumnWidth = 0.83f;
            wsNewExcel.Range["CN1"].ColumnWidth = 6f;
            wsNewExcel.Range["P1"].ColumnWidth = 22f;
            wsNewExcel.Range["Q1"].ColumnWidth = 6f;
            wsNewExcel.Range["AJ1"].ColumnWidth = 10.57f;
            wsNewExcel.Range["AW1"].ColumnWidth = 6f;
            wsNewExcel.Range["BE1"].ColumnWidth = 6f;
            wsNewExcel.Range["BL1"].ColumnWidth = 6f;
            wsNewExcel.Range["BZ1"].ColumnWidth = 6f;
        }

        FileSource endFile = new FileSource() { FileName = "", Path = "" };

        private void AddEndFile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ListViewSelectedEndFile.Items.Count < 1)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog(); //добавление файла
                    openFileDialog.Multiselect = false;
                    openFileDialog.Title = "Выберите документ в который добавятся данные";
                    openFileDialog.Filter = "Таблицы (*.xlsx,*.csv,*.xls)|*.xlsx;*.csv;*.xls"; //форматы файлов, которые отображаются при выборе
                    if (openFileDialog.ShowDialog() == true)
                    {
                        string path = openFileDialog.FileName;
                        string docType = path.Substring(path.Length - 5, 5);
                        Regex regexDocType = new Regex(@"(\.csv)|(\.xlsx)|(\.xls)");
                        if (regexDocType.IsMatch(docType))
                        {
                            if (items.Where(x => x.Path == path).Count() < 1)
                            {
                                ListViewSelectedEndFile.Items.Clear();
                                endFile = new FileSource() { Path = path, FileName = Path.GetFileName(path) };
                                ListViewSelectedEndFile.Items.Add(endFile);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ошибка, можно добавлять только файлы .xlsx, .xls, .csv", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Вы уже добавили файл, в который запишутся данные", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка! Пропробуйте ещё раз.\nКод ошибки: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ListViewSelectedEndFile_PreviewDrop(object sender, DragEventArgs e) //закидываем файл в листвью
        {
            try
            {
                if (ListViewSelectedEndFile.Items.Count < 1)
                {
                    if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    {
                        string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                        if (files.Count() == 1)
                        {
                            string docType = files[0].Substring(files[0].Length - 5, 5);
                            Regex regexDocType = new Regex(@"(\.csv)|(\.xlsx)|(\.xls)");
                            if (regexDocType.IsMatch(docType))
                            {
                                if (items.Where(x => x.Path == files[0]).Count() < 1)
                                {
                                    endFile.Path = files[0];
                                    endFile.FileName = Path.GetFileName(files[0]);
                                    ListViewSelectedEndFile.Items.Clear();
                                    ListViewSelectedEndFile.Items.Add(endFile);
                                }
                            }
                            else
                            {
                                throw new Exception("Ошибка, можно добавлять только файлы .xlsx, .xls, .csv");
                            }

                        }
                        else
                        {
                            throw new Exception("Сюда можно добавить только один файл");
                        }

                    }
                }
                else
                {
                    throw new Exception("Сюда можно добавить только один файл");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ListViewSelectedEndFile_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ListViewSelectedEndFile.SelectedItems.Count > 0)
            {
                DelEndFile.Visibility = Visibility.Visible;
            }
            else
            {
                DelEndFile.Visibility = Visibility.Collapsed;
            }
        }

        private void ListViewSelectedFiles_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ListViewSelectedFiles.SelectedItems.Count > 0)
            {
                DelFiles.Visibility = Visibility.Visible;
            }
            else
            {
                DelFiles.Visibility = Visibility.Collapsed;
            }
        }

        private void DelEndFile_Click(object sender, RoutedEventArgs e)
        {
            endFile.Path = "";
            endFile.FileName = "";
            ListViewSelectedEndFile.Items.Clear();
        }

        private void DelFiles_Click(object sender, RoutedEventArgs e)
        {
            foreach (FileSource file in ListViewSelectedFiles.SelectedItems)
            {
                items.Remove(file);
            }
            ListViewSelectedFiles.Items.Refresh();
        }

        private void StartFiles_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e) //открываем по двойному нажатию
        {
            try
            {
                if (e.ClickCount >= 2)
                {
                    Border b = sender as Border;
                    Process.Start(b.Uid);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EndFile_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if (e.ClickCount >= 2)
                {
                    Process.Start(endFile.Path);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
