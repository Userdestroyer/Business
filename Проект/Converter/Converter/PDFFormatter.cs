using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Threading.Tasks;

class PDFFormatter
{
    static int _sectionNumber = 0;
    static int _pictureNumber = 0;
    static int _tableNumber = 0;
    int _sourceNumber = 0;
    string sourcePath = @"шаблон.txt";
    string distPath = @"result.pdf";
    string[] templateStringList =
    {
         "[*номер раздела*]",
         "[*номер рисунка*]",
         "[*номер таблицы*]",
         "[*ссылка на следующий рисунок*]",
         "[*ссылка на предыдущий рисунок*]",
         "[*ссылка на таблицу*]",
         "[*таблица ",
         "[*cписок литературы*]",
         "[*код",
         "[*рисунок "
    };
    List<string> sourceList = new List<string>();
    public void Make()
    {
        System.IO.FileStream fs = new System.IO.FileStream(distPath, System.IO.FileMode.Create);
        float leftMargin = 50f;
        float rightMargin = 50f;
        Document document = new Document(PageSize.A4, leftMargin, rightMargin, leftMargin,
       rightMargin);
        PdfWriter writer = PdfWriter.GetInstance(document, fs);
        document.Open();
        float fontSizeText = 12f;
        BaseFont baseFont = BaseFont.CreateFont(
        new System.IO.FileInfo(sourcePath).DirectoryName + "\\" + @"ARIAL.TTF",
        BaseFont.IDENTITY_H,
        BaseFont.NOT_EMBEDDED);
        string[] paragraphs = System.IO.File.ReadAllLines(sourcePath);
        foreach (string paragraph in paragraphs)
        {
            bool isSetParagraph = false;

            string textParagraph = paragraph;
            for (int i = 0; i < templateStringList.Length; i++)
            {
                if (paragraph.Contains(templateStringList[i]))
                {
                    switch (i)
                    {
                        case 0:
                            {
                                _sectionNumber++;
                                _pictureNumber = 0;
                                _tableNumber = 0;

                                string replaceString = _sectionNumber.ToString();

                                textParagraph = textParagraph.Replace(templateStringList[i], "");

                                if (_sectionNumber != 1)
                                {
                                    document.NewPage();
                                }

                                var iparagraph = new Paragraph(textParagraph,
                                new Font(baseFont, 13f, Font.BOLD));
                                iparagraph.SpacingAfter = 15f;
                                iparagraph.ExtraParagraphSpace = 10;
                                iparagraph.Alignment = Element.ALIGN_CENTER;
                                Chapter chapter = new Chapter(iparagraph, _sectionNumber);
                                document.Add(chapter);

                                isSetParagraph = true;
                            }
                            break;
                        case 1:
                            {
                                _pictureNumber++;

                                string replaceString = "Рисунок " + _sectionNumber.ToString()
                                + "." + _pictureNumber.ToString() + " –";

                                textParagraph = textParagraph.Replace(templateStringList[i], replaceString);

                                var iparagraph = new Paragraph(textParagraph,
                                new Font(baseFont, fontSizeText, Font.ITALIC));
                                iparagraph.SpacingAfter = 12f;
                                iparagraph.Alignment = Element.ALIGN_CENTER;
                                document.Add(iparagraph);
                                isSetParagraph = true;
                            }
                            break;

                        case 2:
                            {
                                _tableNumber++;
                                string replaceString = "Таблица " + _sectionNumber.ToString()
                                + "." + _tableNumber.ToString() + " –";
                                textParagraph = textParagraph.Replace(templateStringList[i], replaceString);
                                var iparagraph = new Paragraph(textParagraph,
                                new Font(baseFont, fontSizeText, Font.ITALIC));
                                iparagraph.SpacingAfter = 12f;
                                iparagraph.Alignment = Element.ALIGN_LEFT;
                                document.Add(iparagraph);

                                isSetParagraph = true;
                            }
                            break;

                        case 3:
                            {
                                string replaceString = _sectionNumber.ToString() + "." + (_pictureNumber + 1).ToString();
                                textParagraph = textParagraph.Replace(templateStringList[i], replaceString);
                            }
                            break;
                        case 4:
                            {
                                string replaceString = _sectionNumber.ToString() + "." + _pictureNumber.ToString();
                                textParagraph = textParagraph.Replace(templateStringList[i], replaceString);
                            }
                            break;
                        case 5:
                            {
                                string replaceString = _sectionNumber.ToString() + "." + (_tableNumber + 1).ToString();
                                textParagraph = textParagraph.Replace(templateStringList[i], replaceString);
                            }
                            break;

                        case 6:
                            {
                                string csvPath = textParagraph.Replace(templateStringList[i], "")
                                .Replace("*", "").Replace("\r", "").Replace("]", "");

                                csvPath = new System.IO.FileInfo(sourcePath).DirectoryName + "\\" + csvPath;

                                string[] listRows = System.IO.File.ReadAllLines(csvPath);

                                string[] listTitle = listRows[0].Split(";,".ToCharArray(),
                                StringSplitOptions.RemoveEmptyEntries);

                                PdfPTable table = new PdfPTable(listTitle.Length);

                                foreach (string title in listTitle)
                                {
                                    PdfPCell cell = new PdfPCell(new Phrase(title.ToString(),
                                            new Font(baseFont, fontSizeText, Font.BOLDITALIC, Color.YELLOW)));
                                    cell.BackgroundColor = Color.GRAY;
                                    cell.BorderWidthBottom = 1;
                                    table.AddCell(cell);
                                }
                                int bgColorIterator = 0;
                                foreach (string row in listRows)
                                {
                                    bgColorIterator++;
                                    if (row == listRows[0]) continue;

                                    string[] listValue = row.Split(";,".ToCharArray(),
                                            StringSplitOptions.RemoveEmptyEntries);
                                    foreach (string value in listValue)
                                    {
                                        PdfPCell style = new PdfPCell();
                                        style.BorderColor = Color.RED;
                                        style.HorizontalAlignment = 1;

                                        PdfPCell cell = new PdfPCell(new Phrase(value.ToString(),
                                                new Font(baseFont, fontSizeText, Font.HELVETICA, Color.CYAN)));
                                        cell.Border = 0;

                                        if (bgColorIterator % 2 == 0)
                                            cell.BackgroundColor = Color.GRAY;
                                        else
                                            cell.BackgroundColor = Color.WHITE;

                                        table.AddCell(cell);
                                    }
                                }

                                document.Add(table);

                                isSetParagraph = true;
                            }
                            break;

                        case 7:
                            {
                                string replaceString = "";
                                for (int j = 0; j < sourceList.Count; j++)
                                {
                                    replaceString = (j + 1).ToString() + ". "
                                    + sourceList[j].TrimStart('[').TrimEnd(']') + "\r\n";

                                    var iparagraph = new Paragraph(replaceString,
                                    new Font(baseFont, fontSizeText, Font.NORMAL));
                                    iparagraph.SpacingAfter = 0;
                                    iparagraph.SpacingBefore = 0;
                                    iparagraph.FirstLineIndent = 20f;
                                    iparagraph.ExtraParagraphSpace = 10;
                                    iparagraph.Alignment = Element.ALIGN_JUSTIFIED;
                                    document.Add(iparagraph);
                                }

                                isSetParagraph = true;
                            }
                            break;

                        case 8:
                            {
                                //если есть шаблонная строка для места вставки кода
                                textParagraph = "тут будет ваш код";
                                //TODO (задание на 5) вставить код из файла - CourierNew 8 пт одинарный без отступа в рамке
                            }
                            break;

                        case 9:
                            {
                                string jpgPath = textParagraph.Replace(templateStringList[i],
                                 "").Replace("*", "").Replace("\r", "").Replace("]", "");

                                 jpgPath = new System.IO.FileInfo(sourcePath).DirectoryName
                                 + "\\" + jpgPath;

                                Image jpg = Image.GetInstance(jpgPath);
                                jpg.Alignment = Element.ALIGN_CENTER;
                                jpg.SpacingBefore = 12f;

                                float procent = 90;
                                while (jpg.ScaledWidth > PageSize.A4.Width / 2.0f)
                                {
                                    jpg.ScalePercent(procent);
                                    procent -= 10;
                                }

                                document.Add(jpg);

                                isSetParagraph = true;
                            }
                            break;
                    }
                }
            }

            string text = textParagraph;

            if (text.Contains("["))
            {
                for (int j = 0; j < text.Length - 1; j++)
                {
                    if (text[j] == '[' && text[j + 1] != '*')
                    {
                        int startIndex = j;
                        int endIndex = startIndex + 1;
                        while (endIndex < text.Length
                        &&
                       text[endIndex] != ']')
                        {
                            endIndex++;
                        }
                        string sourceName = "";

                        if (text[endIndex] == ']')
                        {
                            for (int k = startIndex; k <= endIndex; k++)
                            {
                                sourceName += text[k];
                            }
                            int index = 0;

                            if (!sourceList.Contains(sourceName))
                            {
                                sourceList.Add(sourceName);
                                _sourceNumber++;
                                index = _sourceNumber;
                            }
                            else
                            {
                                for (int k = 0; k < sourceList.Count; k++)
                                {

                                    if (sourceList[k].Contains(sourceName))
                                    {
                                        index = k + 1;
                                    }
                                }
                            }

                            string replaceString = "[" + index.ToString() + "]";

                            textParagraph = textParagraph.Replace(sourceName, replaceString);

                            textParagraph = textParagraph.Replace(' ' + sourceName,
                                            "\u00A0[" + index.ToString() + "]");

                            j = endIndex;
                        }
                    }
                }
            }

            textParagraph = textParagraph.Replace("  ", " ");

            if (!isSetParagraph)
            {

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

        document.Close();
    }
}