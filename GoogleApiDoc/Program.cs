using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System;
using System.Reflection;
using Microsoft.Office.Interop.Word;


namespace GoogleApiDoc
{
    internal class Program
    //  {
    //    static void Main(string[] args)
    //  {
    //}
    //}
    //}
    {
        /// <summary>
        /// номера раздела, ==0 - нет разделов
        /// </summary>
        static uint _sectionNumber = 0;
        /// <summary>
        /// номера рисунка, ==0 - нет картинок
        /// </summary>
        static uint _pictureNumber = 0;
        /// <summary>
        /// номера таблиц, ==0 - нет таблиц
        /// </summary>
        static uint _tableNumber = 0;

        string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        static void Main(string[] args)
        {
            //string sourcePath = @"D:\Repos\SunstriderZ37\Файлы\шаблон.rtf";//путь до исходного шаблона
            string sourcePath = @"c:\Users\pc-rzn.ru\Dropbox\ПК\Documents\GitHub\GoogleApiDoc\шаблон.rtf";//путь до исходного шаблона
            //string distPath = @"D:\Repos\SunstriderZ37\Файлы\result.rtf";//путь до выходного файла
            string distPath = @"c:\Users\pc-rzn.ru\Dropbox\ПК\Documents\GitHub\GoogleApiDoc\result.pdf";//путь до выходного файла
            //string csvPath = @"D:\Repos\SunstriderZ37\Файлы\data.csv";//путь до csv файла для создания таблицы
            string csvPath = @"c:\Users\pc-rzn.ru\Dropbox\ПК\Documents\GitHub\GoogleApiDoc\data.csv";//путь до csv файла для создания таблицы
            //string codePath = @"D:\Repos\SunstriderZ37\AltukhovZ37\AltukhovZ37\program.cs";//путь до файла с кодом
            string codePath = @"c:\Users\pc-rzn.ru\Dropbox\ПК\Documents\GitHub\GoogleApiDoc\program.cs";//путь до файла с кодом

            //список закладок
            string[] templateStringList =
                {
                "[*имя раздела*]",///0
                "[*имя рисунка*]",///1
                "[*ссылка на следующий рисунок*]",///2
                "[*ссылка на предыдущий рисунок*]",///3
                "[*ссылка на таблицу*]",///4
                "[*таблица первая*]",///5
                "[*имя таблицы*]",///6
                "[*код*]"///7
                };

            var application = new Application();
            application.Visible = true;

            var document = application.Documents.Open(sourcePath, false);

            Paragraph prevParagraph = null;
            Object missing = System.Type.Missing;

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                for (int i = 0; i < templateStringList.Length; i++)
                {
                    if (paragraph.Range.Text.Contains(templateStringList[i]))
                    {
                        switch (i)
                        {
                            case 0:
                                {
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; //Выравниевание по центру
                                    paragraph.Range.Font.Name = "Times New Roman"; //Шрифт
                                    paragraph.Range.Font.Size = 15; //Размер текста
                                    paragraph.Format.SpaceAfter = 12; //Отступ
                                    paragraph.Range.Font.Bold = 1;//Стиль шрифта
                                    paragraph.Range.HighlightColorIndex = 0;//Выделение текста

                                    _sectionNumber++;
                                    _pictureNumber = 0;
                                    string replaceString = _sectionNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                  ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                   0, ref missing, replaceString, 2, ref missing, ref missing,
                                  ref missing, ref missing);
                                }
                                break;
                            case 1:
                                {
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; //Выравниевание по центру
                                    paragraph.Range.Font.Name = "Times New Roman"; //Шрифт
                                    paragraph.Range.Font.Size = 12; //Размер текста
                                    paragraph.Format.SpaceAfter = 12; //Отступ
                                    paragraph.Range.HighlightColorIndex = 0; //Выделение текста

                                    if (prevParagraph != null)
                                    {
                                        prevParagraph.Format.SpaceBefore = 12; //Отступ
                                        prevParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; //Выравниевание по центру
                                    }

                                    _pictureNumber++;
                                    string replaceString = "Рисунок " + _sectionNumber.ToString() + "." + _pictureNumber.ToString() + " -";

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 2:
                                {
                                    paragraph.Range.HighlightColorIndex = 0;
                                    string replaceString = _sectionNumber.ToString() + "." + (_pictureNumber + 1).ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 3:
                                {
                                    paragraph.Range.HighlightColorIndex = 0;
                                    string replaceString = _sectionNumber.ToString() + "." + _pictureNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 4:
                                {
                                    paragraph.Range.HighlightColorIndex = 0;
                                    _tableNumber++;
                                    string replaceString = _sectionNumber.ToString() + "." + _tableNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                    0, ref missing, replaceString, 2, ref missing, ref missing,
                                   ref missing, ref missing);
                                }
                                break;
                            case 5:
                                {
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    var range = application.Selection.Range;
                                    range.HighlightColorIndex = 0;


                                    string[] listRows = System.IO.File.ReadAllText(csvPath).Split("\r\n".ToCharArray(), StringSplitOptions.
                                                                                    RemoveEmptyEntries);

                                    string[] listTitle = listRows[0].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                    var wordTable = document.Tables.Add(range, listRows.Length, listTitle.Length);

                                    wordTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle; //Внутренние границы таблицы
                                    wordTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle; //Внешние границы таблицы

                                    for (var k = 0; k < listTitle.Length; k++)
                                    {
                                        wordTable.Cell(1, k + 1).Range.Text = listTitle[k].ToString();
                                    }
                                    for (var j = 1; j < listRows.Length; j++)
                                    {
                                        string[] listValues = listRows[j].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                                        for (var k = 0; k < listValues.Length; k++)
                                        {
                                            wordTable.Cell(j + 1, k + 1).Range.Text = listValues[k].ToString();
                                        }
                                    }
                                }
                                break;
                            case 6:
                                {
                                    {
                                        paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphLeft; //Выравнивание текста по левому краю
                                        paragraph.Range.Font.Name = "Times New Roman"; //Шрифт
                                        paragraph.Range.Font.Size = 14; //Размер текста
                                        paragraph.Range.HighlightColorIndex = 0; //Выделение текста

                                        _tableNumber++;
                                        string replaceString = "Таблица " + _sectionNumber.ToString() + "." + _tableNumber.ToString() + " -";

                                        paragraph.Range.Find.Execute(templateStringList[i],
                                       ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                        0, ref missing, replaceString, 2, ref missing, ref missing,
                                       ref missing, ref missing);
                                    }
                                }
                                break;
                            case 7:
                                {
                                    application.Selection.Find.Execute(templateStringList[i]);
                                    var range = application.Selection.Range;
                                    range.HighlightColorIndex = 0;

                                    string code = System.IO.File.ReadAllText(codePath);
                                    range.Text = code;
                                }
                                break;
                        }
                    }
                }
                prevParagraph = paragraph;
            }

            document.SaveAs2(distPath);
            System.Console.In.Read();
            application.Quit();
        }
    }
}
