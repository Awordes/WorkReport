using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OfficeOpenXml;
using Xceed.Wpf.Toolkit;
using System.IO;
using Microsoft.Win32;

namespace WorkReport
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //1 семестр:
        //			лекции, практические, курсовые - сентябрь, октябрь, ноябрь, декабрь
        //          контрольные работы - ноябрь
        //			зачеты - декабрь
        //			экзамены, консультации - январь
        //2 семестр:
        //			лекции, практические, курсовые - февраль, март, апрель, май
        //          контрольные работы - апрель
        //			зачеты, экзамены, консультации - июнь
        //руководство ВКР - февраль, март, апрель, май
        //допуск - июнь
        //практика - июнь-июль
        //руководство кафедрой - сентябрь, октябрь, ноябрь, декабрь, январь, февраль, март, апрель, май
        //Аспиранты - сентябрь, октябрь, ноябрь, декабрь, январь, февраль, март, апрель, май

        //Номера листов отчета (начиная с 1)
        int September = 1;
        int October = 2;
        int November = 3;
        int December = 4;
        int January = 5;
        int February = 6;
        int March = 7;
        int April = 8;
        int May = 9;
        int June = 10;
        int Year = 11;
        int[] year = new int[11];
        int MonthCount = 4; //Количество рабочий месяцев в семестре

        //Нагрузка преподавателя ("А" = 1)
        int[] TeachWorkOld = new int[22];
        string[] TeachWork = new string[22]; //столбцы часов

        int Semester = 4; //столбец семестра
        int RowStart = 6; //номер первой строки с предметом

        //Отчет по нагрузке ("А" = 1)
        int[] ReportWorkOld = new int[17];
        string[] ReportWork = new string[17];
        int RowNumber = 0; //номер строки с преподавателем

        string PathIn = ""; //путь к файлу с нагрузкой преподавателя
        string PathOut = ""; //путь к отчету с нагрузкой

        ExcelPackage InputFile = null; //файл с нагрузкой преподавателя
        ExcelPackage OutputFile = null; //отчет с нагрузкой

        string TeacherName = ""; //Имя преподавателя

        public MainWindow()
        {
            InitializeComponent();

            year[0] = 1;
            year[1] = 2;
            year[2] = 3;
            year[3] = 4;
            year[4] = 5;
            year[5] = 6;
            year[6] = 7;
            year[7] = 8;
            year[8] = 9;
            year[9] = 10;
            year[10] = 11;

            TeachWork[0] = "I"; //Лекции
            TeachWork[1] = "J"; //Практические занятия
            TeachWork[2] = "K"; //Лабораторные занятия
            TeachWork[3] = "L"; //Зачеты
            TeachWork[4] = "M"; //Экзамены
            TeachWork[5] = "N"; //Консультации экзаменов
            TeachWork[6] = "O"; //Консультации зачетов
            TeachWork[7] = "P"; //Контрольные работы
            TeachWork[8] = "Q"; //Курсовые работы
            TeachWork[9] = "R"; //Курсовые проекты
            TeachWork[10] = "S"; //РГР
            TeachWork[11] = "T"; //Производственные практики
            TeachWork[12] = "U"; //Учебные практики
            TeachWork[13] = "W"; //Преддипломные практики
            TeachWork[14] = "X"; //НИРС
            TeachWork[15] = "Y"; //ГЭК
            TeachWork[16] = "Z"; //Руководство и консультирование
            TeachWork[17] = "AA"; //Рецензирование
            TeachWork[18] = "AC"; //Допуск к защите
            TeachWork[19] = "AD"; //Защита ВКР
            TeachWork[20] = "AE"; //Руководство кафедрой
            TeachWork[21] = "AF"; //Другая нагрузка

            ReportWork[0] = "A"; //Номер
            ReportWork[1] = "B"; //ФИО
            ReportWork[2] = "C"; //Лекции
            ReportWork[3] = "D"; //Практические/лабораторные занятия
            ReportWork[4] = "E"; //Консультации
            ReportWork[5] = "F"; //Контрольные работы
            ReportWork[6] = "G"; //Расчетно-графические работы
            ReportWork[7] = "H"; //Курсовые работы/проекты
            ReportWork[8] = "I"; //Зачет
            ReportWork[9] = "J"; //Экзамен
            ReportWork[10] = "K"; //Учебная/производственная практика
            ReportWork[11] = "L"; //Преддипломная практика
            ReportWork[12] = "M"; //НИРС
            ReportWork[13] = "N"; //ВКР
            ReportWork[14] = "O"; //ГЭК
            ReportWork[15] = "P"; //Аспиранты/докторанты
            ReportWork[16] = "Q"; //Другие виды работ
        }

        private void buttonIn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfd = new OpenFileDialog();
            openfd.Filter = "Excel Files|*.xlsx";
            openfd.Title = "Открыть...";
            if (openfd.ShowDialog() == true)
            {
                PathIn = openfd.FileName;
                textBoxIn.Text = PathIn;
            }
        }

        private void buttonOut_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfd = new OpenFileDialog();
            openfd.Filter = "Excel Files|*.xlsx";
            openfd.Title = "Открыть...";
            if (openfd.ShowDialog() == true)
            {
                PathOut = openfd.FileName;
                textBoxOut.Text = PathOut;
            }
        }

        private void buttonCalculate_Click(object sender, RoutedEventArgs e)
        {
            RowNumber = Convert.ToInt32(nUDRowNumber.Value);
            try
            {
                InputFile = new ExcelPackage(new FileInfo(PathIn));
                OutputFile = new ExcelPackage(new FileInfo(PathOut));
            }
            catch
            {
                System.Windows.MessageBox.Show("Файл используется другим приложением");
                return;
            }
            ExcelWorksheet inPage = InputFile.Workbook.Worksheets[1];
            TeacherName = inPage.Cells[3, 1].Value.ToString();
            int i = RowStart;

            //обнуление значений таблиц
            for (int k = 0; k <= 9; k++)
            {
                ExcelWorksheet Month = OutputFile.Workbook.Worksheets[year[k]];
                for (int l = 2; l < ReportWork.Length; l++)
                    Month.Cells[ReportWork[l] + RowNumber.ToString()].Value = 0;
            }

            do
            {
                int start = 0;
                int end = 9;
                int special = 10;
                int sem = inPage.Cells[i, Semester].Value == null ||
                    inPage.Cells[i, Semester].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[i, Semester].Value);
                if (sem == 1)
                {
                    start = 0;
                    end = 3;
                    special = 4;
                }
                else if (sem == 2)
                {
                    start = 5;
                    end = 8;
                    special = 9;
                }
                else if (sem == 0)
                {
                    start = 0;
                    end = 9;
                    special = 9;
                }
                int value = 0;

                //Лекции
                value = inPage.Cells[TeachWork[0] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[0] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[0] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 2);

                //Практические работы
                value = inPage.Cells[TeachWork[1] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[1] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[1] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 3);

                //Лабораторные работы
                value = inPage.Cells[TeachWork[2] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[2] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[2] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 3);

                //Зачеты
                value = inPage.Cells[TeachWork[3] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[3] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[3] + i.ToString()].Value);
                if (sem == 1) AddNewHourToPrevious(end, end, value, 8);
                else AddNewHourToPrevious(special, special, value, 8);

                //Экзамены
                value = inPage.Cells[TeachWork[4] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[4] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[4] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 9);

                //Консультации экзаменов
                value = inPage.Cells[TeachWork[5] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[5] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[5] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 4);

                //Консультации зачетов
                value = inPage.Cells[TeachWork[6] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[6] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[6] + i.ToString()].Value);
                if (sem == 1) AddNewHourToPrevious(end, end, value, 4);
                else AddNewHourToPrevious(special, special, value, 4);

                //Контрольные работы
                value = inPage.Cells[TeachWork[7] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[7] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[7] + i.ToString()].Value);
                AddNewHourToPrevious(end - 1, end - 1, value, 5);

                //Курсовые работы
                value = inPage.Cells[TeachWork[8] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[8] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[8] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 7);

                //Курсовые проекты
                value = inPage.Cells[TeachWork[9] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[9] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[9] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 7);

                //РГР
                value = inPage.Cells[TeachWork[10] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[10] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[10] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 6);

                //Производственные практики
                value = inPage.Cells[TeachWork[11] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[11] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[11] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 10);

                //Учебные практики
                value = inPage.Cells[TeachWork[12] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[12] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[12] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 10);

                //Преддипломные практики
                string ind = TeachWork[13] + i.ToString();
                value = inPage.Cells[TeachWork[13] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[13] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[13] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 11);

                //НИРС
                value = inPage.Cells[TeachWork[14] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[14] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[14] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 12);

                //ГЭК
                value = inPage.Cells[TeachWork[15] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[15] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[15] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 14);

                //Руководство и консультирование
                value = inPage.Cells[TeachWork[16] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[16] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[16] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 13);

                //Рецензирование
                value = inPage.Cells[TeachWork[17] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[17] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[17] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 13);

                //Допуск к защите
                value = inPage.Cells[TeachWork[18] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[18] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[18] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 13);

                //Защита ВКР
                value = inPage.Cells[TeachWork[19] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[19] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[19] + i.ToString()].Value);
                AddNewHourToPrevious(start, end, value, 13);

                //Руководство кафедрой
                value = inPage.Cells[TeachWork[20] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[20] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[20] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 16);

                //Допуск к защите
                value = inPage.Cells[TeachWork[21] + i.ToString()].Value == null ||
                    inPage.Cells[TeachWork[21] + i.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(inPage.Cells[TeachWork[21] + i.ToString()].Value);
                AddNewHourToPrevious(special, special, value, 16);

                i++;
            }
            while (inPage.Cells[i, 2].Value.ToString() != "ВСЕГО");
            System.Windows.MessageBox.Show("Done");
            OutputFile.Save();
        }


        void AddNewHourToPrevious(int start, int end, int value, int reportIndex)
        {
            for (int k = start; k <= end; k++)
            {
                ExcelWorksheet Month = OutputFile.Workbook.Worksheets[year[k]];

                int temp = Month.Cells[ReportWork[reportIndex] + RowNumber.ToString()].Value == null ||
                    Month.Cells[ReportWork[reportIndex] + RowNumber.ToString()].Value.ToString() == ""
                    ? 0
                    : Convert.ToInt32(Month.Cells[ReportWork[reportIndex] + RowNumber.ToString()].Value);
                if (start == end) temp += value;
                else temp += value / (end - start + 1);
                if (k == end && start != end) temp += value % (end - start + 1);
                Month.Cells[ReportWork[reportIndex] + RowNumber.ToString()].Value = temp;
            }
        }

        private void buttonAbout_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("by Awordes"+ Environment.NewLine + 
                                            "Andrew Tolstov" + Environment.NewLine + 
                                            "awordes76@gmail.com" + Environment.NewLine +
                                            "https://github.com/Awordes/WorkReport");
        }
    }
}
