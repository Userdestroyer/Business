﻿using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;

/// <summary>класс для формирования документа PDF</summary>
class PDFFormatter
{
    /// <summary>текущий номер раздела в тексте</summary>
    static int _sectionNumber = 0;
    /// <summary>текущий номер рисунка в тексте<</summary>
    static int _pictureNumber = 0;
    /// <summary>текущий номер таблиц в тексте<</summary>
    static int _tableNumber = 0;
    /// <summary>нумерация источников в списке литературы</summary>
    int _sourceNumber = 0;
    /// <summary>путь до исходного шаблона</summary>
    string sourcePath = @"C:\Users\user\YandexDisk\ТУСУР\5 курс\2 семестр\СПРБП\Business\Исходники\шаблон.txt";
    /// <summary>путь до выходного файла</summary>
    string distPath = @"C:\Users\user\YandexDisk\ТУСУР\5 курс\2 семестр\СПРБП\Business\Результат\result.pdf";
    /// <summary>список шаблонных строк в тексте для форматирования</summary>
    string[] templateStringList =
    {
         "[*номер раздела*]", //0
         "[*номер рисунка*]", //1
         "[*номер таблицы*]", //2
         "[*ссылка на следующий рисунок*]", //3
         "[*ссылка на предыдущий рисунок*]", //4
         "[*ссылка на таблицу*]", //5
         "[*таблица ", //6
         "[*cписок литературы*]", //7
         "[*код", //8
         "[*рисунок " //9
    };
    /// <summary>список литературы</summary>
    List<string> sourceList = new List<string>();
    public void Make()
    {
        //CODEPART 1 Открытие документа и задание его формата, подготовка
        System.IO.FileStream fs = new System.IO.FileStream(distPath, System.IO.FileMode.Create);
        //обозначаем размер полей
        float leftMargin = 50f;
        float rightMargin = 50f;
        //создаем документ
        Document document = new Document(PageSize.A4, leftMargin, rightMargin, leftMargin,
       rightMargin);
        //связываем документ с файлом
        PdfWriter writer = PdfWriter.GetInstance(document, fs);
        //открываем документ
        document.Open();
        //определяем шрифт
        float fontSizeText = 12f;
        BaseFont baseFont = BaseFont.CreateFont(
        new System.IO.FileInfo(sourcePath).DirectoryName + "\\" + @"ARIAL.TTF",
        BaseFont.IDENTITY_H,
        BaseFont.NOT_EMBEDDED);
        //считываем все строки из текстового файла
        string[] paragraphs = System.IO.File.ReadAllLines(sourcePath);
        //CODEPART 2 обходим все строки файла - параграфы
        foreach (string paragraph in paragraphs)
        {
            //вставлен ли уже параграф в PDF
            bool isSetParagraph = false;
            //текущий текст параграфа
            string textParagraph = paragraph;
            //проверяем, входит ли в параграф ключевое слово
            for (int i = 0; i < templateStringList.Length; i++)
            {
                if (paragraph.Contains(templateStringList[i]))
                {
                    switch (i)
                    {
                        //CODEPART 2.1 Редактирование абзаца заголовка раздела
                        case 0:// "[*номер раздела*]"
                            {
                            }
                            break;
                        //CODEPART 2.1 Редактирование подрисуночной подписи
                        case 1://"[*номер рисунка*]"
                            {
                            }
                            break;
                        //CODEPART 2.3 Редактирование заголовка таблицы
                        case 2://"[*номер таблицы*]"
                            {
                            }
                            break;

                        //CODEPART 2.4 Вставка ссылки на следующий рисунок
                        case 3://"[*ссылка на следующий рисунок*]"
                            {
                            }
                            break;
                        //CODEPART 2.5 Вставка ссылки на предыдущий рисунок
                        case 4://"[*ссылка на таблицу*]
                            {
                            }
                            break;
                        //CODEPART 2.6 Вставка ссылки на таблицу
                        case 5://"[*ссылка на таблицу*]"
                            {
                            }
                            break;
                        //CODEPART 2.7 Вставка таблицы из файла
                        case 6://"[*таблица "
                            {
                            }
                            break;
                        //CODEPART 2.8 Вставка списка литературы
                        case 7://"[*cписок литературы*]"
                            {
                            }
                            break;
                        //CODEPART 2.9 Вставка кода из файла
                        case 8://"[*код"
                            {
                            }
                            break;
                        //CODEPART 2.10 Вставка рисунка из файл
                        case 9://"[*таблица "
                            {
                            }
                            break;
                    }
                }
            }
            //CODEPART 2.11 Сбор внутритекстовых ссылок на литературу
            //CODEPART 2.12 Стандартное форматирование абзаца
            //если нужно абзац форматировать как обычный текст и абзац еще не вставлен
            if (!isSetParagraph)
            {
                //вставляем абзац со стандарным форматированием
                var iparagraph = new Paragraph(textParagraph,
                 new Font(baseFont, fontSizeText, Font.NORMAL));
                iparagraph.SpacingAfter = 0;
                iparagraph.SpacingBefore = 0;
                iparagraph.FirstLineIndent = 20f;
                iparagraph.ExtraParagraphSpace = 10;
                iparagraph.Alignment = Element.ALIGN_JUSTIFIED;
                document.Add(iparagraph);
            }
        }
        //закрываем документ
        document.Close();
    }
}