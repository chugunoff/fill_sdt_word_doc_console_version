using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;
using BlipFill = DocumentFormat.OpenXml.Drawing.Pictures.BlipFill;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.Drawing;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using DocumentFormat.OpenXml.Office2013.Word;

namespace OpenXMLFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            String pathFile = @"C:\Users\chugu\Desktop\Projects\FillWordDoc — копия.docx";
            String xml =
                @"<?xml version=""1.0"" encoding=""utf-8"" ?>
<macroVars>
    <richText name=""ДокЗаголовок""><![CDATA[PNCd0L7QstGL0Lkg0LfQsNCz0L7Qu9C+0LLQvtC6INC00L7QutGD0LzQtdC90YLQsD4=]]></richText>
    <richText name=""ПустойФорматТекст""><![CDATA[0J/QtdGA0LLRi9C5INCw0LHQt9Cw0YbQuNC6Lg0K0JLRgtC+0YDQvtC5INCw0LHQt9Cw0YbQuNC6Lg==]]></richText>
    <table name=""НормальнаяТаблица""><![CDATA[0Y/Rh9C10LnQutCwMTENCtGP0YfQtdC50LrQsDExDQrRj9GH0LXQudC60LAxMV5eXtGP0YfQtdC50LrQsDEyDQrRj9GH0LXQudC60LAxMg0K0Y/Rh9C10LnQutCwMTJeXl7Rj9GH0LXQudC60LAxMw0K0Y/Rh9C10LnQutCwMTMNCtGP0YfQtdC50LrQsDEzXl5e0Y/Rh9C10LnQutCwMTQNCtGP0YfQtdC50LrQsDE0DQrRj9GH0LXQudC60LAxNHx8fNGP0YfQtdC50LrQsDIxXl5e0Y/Rh9C10LnQutCwMjJeXl7Rj9GH0LXQudC60LAyM15eXtGP0YfQtdC50LrQsDI0DQo=]]></table>
    <picture name=""Рисунок""><![CDATA[QzpcVXNlcnNcYXUwMDAxMzZcRGVza3RvcFxBTEYuanBn]]></picture>
    <repeatedSection name=""ПовторРаздел"">
        <repeatedSectionItem id=""0"">
            <richText name=""ТекстВРазделе""><![CDATA[0KLQtdC60YHRgiDQsiDRgNCw0LfQtNC10LvQtQ==]]></richText>
            <richText name=""ТекстВРазделе_2""><![CDATA[0KLQtdC60YHRgiDQsiDRgNCw0LfQtNC10LvQtV8y]]></richText>
            <repeatedSection name=""Подраздел"">
                <repeatedSectionItem id=""0"">
                    <richText name=""ТекстПодраздел""><![CDATA[0KLQtdC60YHRgtCf0L7QtNGA0LDQt9C00LXQu18x]]></richText>
                </repeatedSectionItem>
                <repeatedSectionItem id=""1"">
                    <richText name=""ТекстПодраздел""><![CDATA[0KLQtdC60YHRgtCf0L7QtNGA0LDQt9C00LXQu18y]]></richText>
                </repeatedSectionItem>
                <repeatedSectionItem id=""2"">
                    <richText name=""ТекстПодраздел""><![CDATA[0KLQtdC60YHRgtCf0L7QtNGA0LDQt9C00LXQu18z]]></richText>
                </repeatedSectionItem>
            </repeatedSection>
        </repeatedSectionItem>
        <repeatedSectionItem id=""1"">
            <richText name=""ТекстВРазделе""><![CDATA[0KLQtdC60YHRgiDQsiDRgNCw0LfQtNC10LvQtSDRgdC10LrRhtC40LggMg==]]></richText>
            <richText name=""ТекстВРазделе_2""><![CDATA[0KLQtdC60YHRgiDQsiDRgNCw0LfQtNC10LvQtSDRgdC10LrRhtC40LggMl8y]]></richText>
            <repeatedSection name=""Подраздел"">
                <repeatedSectionItem id=""0"">
                    <richText name=""ТекстПодраздел""><![CDATA[0KLQtdC60YHRgtCf0L7QtNGA0LDQt9C00LXQu18x]]></richText>
                </repeatedSectionItem>
                <repeatedSectionItem id=""1"">
                    <richText name=""ТекстПодраздел""><![CDATA[0KLQtdC60YHRgtCf0L7QtNGA0LDQt9C00LXQu18y]]></richText>
                </repeatedSectionItem>
                <repeatedSectionItem id=""2"">
                    <richText name=""ТекстПодраздел""><![CDATA[0KLQtdC60YHRgtCf0L7QtNGA0LDQt9C00LXQu18z]]></richText>
                </repeatedSectionItem>
            </repeatedSection>
        </repeatedSectionItem>
    </repeatedSection>
</macroVars>";

            // используем using вместо .Open .Save .Close
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(pathFile, true))
            {
                // Спарсить полученный xml
                XmlDocument xDoc = new XmlDocument();
                xDoc.LoadXml(xml);

                // Получить корневой элемент
                XmlElement macroList = xDoc.DocumentElement;

                // Получить тело документа
                Body body = wdDoc.MainDocumentPart.Document.Body;

                Fill(body, macroList, wdDoc);

                Console.ReadKey();
            }
        }

        public static void Fill(OpenXmlElement docNode, XmlElement macroListXml, WordprocessingDocument wdDoc)
        {
            /* Форматированный текст для встроенных элементов */
            IEnumerable<Paragraph> paragraphs = docNode.Elements<Paragraph>();
            foreach (Paragraph paragraph in paragraphs)
            {
                IEnumerable<SdtRun> sdtRuns = paragraph.Elements<SdtRun>();
                foreach (SdtRun sdtRun in sdtRuns)
                {
                    // Из SdtProperties взять Tag для идентификации Content Control
                    SdtProperties sdtProperties = sdtRun.GetFirstChild<SdtProperties>();
                    Tag tag = sdtProperties.GetFirstChild<Tag>();

                    // Найти в macroListXml Node макропеременной
                    String macroVarValue = FindMacroVar(macroListXml, tag.Val);

                    if (macroVarValue != null)
                    {
                        // Сохранить старый стиль Run
                        SdtContentRun sdtContentRun = sdtRun.GetFirstChild<SdtContentRun>();
                        OpenXmlElement oldRunProps = sdtContentRun.GetFirstChild<Run>().GetFirstChild<RunProperties>().CloneNode(true);

                        // Очистить Node Content Control
                        sdtContentRun.RemoveAllChildren();

                        // Создать новую Run Node
                        Run newRun = sdtContentRun.AppendChild(new Run());
                        // Вернуть старый стиль
                        newRun.AppendChild(oldRunProps);

                        // Вставить текст (без переносов строк!!!)
                        newRun.AppendChild(new Text(macroVarValue));
                    }
                }
            }

            /* Получить остальные Content Control */
            IEnumerable<SdtBlock> sdtBlocks = docNode.Elements<SdtBlock>();
            foreach (SdtBlock sdtBlock in sdtBlocks)
            {
                // Получить параметры(SdtProperties) Content Control
                SdtProperties sdtProperties = sdtBlock.GetFirstChild<SdtProperties>();

                // Получить Tag для идентификации Content Control
                Tag tag = sdtProperties.GetFirstChild<Tag>();

                // Получить значение макроперенной из macroListXml
                Console.WriteLine("Tag: " + tag.Val);
                String macroVarValue = FindMacroVar(macroListXml, tag.Val);

                // Если макропеременная есть в MacroListXml
                if (macroVarValue != null)
                {
                    Console.WriteLine("Value: " + macroVarValue);
                    // Получить блок содержимого Content Control
                    SdtContentBlock sdtContentBlock = sdtBlock.GetFirstChild<SdtContentBlock>();

                    /* Форматированный текст для абзацев */
                    if (sdtProperties.GetFirstChild<SdtPlaceholder>() != null && sdtContentBlock.GetFirstChild<Paragraph>() != null)
                    {
                        // Сохранить старый стиль параграфа
                        ParagraphProperties oldParagraphProperties = sdtContentBlock.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().CloneNode(true) as ParagraphProperties;
                        String oldParagraphPropertiesXml = oldParagraphProperties.InnerXml;

                        // Очистить ноду с контентом
                        sdtContentBlock.RemoveAllChildren();

                        InsertText(macroVarValue, oldParagraphPropertiesXml, sdtContentBlock);
                    }

                    /* Таблицы */
                    if (sdtProperties.GetFirstChild<SdtPlaceholder>() != null && sdtContentBlock.GetFirstChild<Table>() != null)
                    {
                        // Получить ноду таблицы
                        Table table = sdtContentBlock.GetFirstChild<Table>();

                        // Получить все строки таблицы
                        IEnumerable<TableRow> tableRows = table.Elements<TableRow>();

                        // Получить вторую строку из таблицы
                        TableRow tableRow = tableRows.ElementAt(1) as TableRow;
                        Type tableRowType = tableRow.GetType();

                        // Получить все стили столбцов
                        List<String> paragraphCellStyles = new List<string>();
                        IEnumerable<OpenXmlElement> tableCells = tableRow.Elements<TableCell>();
                        foreach (OpenXmlElement tableCell in tableCells)
                        {
                            String paragraphCellStyleXml = tableCell.GetFirstChild<Paragraph>().GetFirstChild<ParagraphProperties>().InnerXml;
                            paragraphCellStyles.Add(paragraphCellStyleXml);
                        }

                        // Удалить все строки, после первой
                        while (tableRows.Count<TableRow>() > 1)
                        {
                            TableRow lastTableRows = tableRows.Last<TableRow>();
                            lastTableRows.Remove();
                        }

                        // Удалить последний элемент, если это не TableRow
                        OpenXmlElement lastNode = table.LastChild;
                        if (lastNode.GetType() != tableRowType)
                        {
                            lastNode.Remove();
                        }

                        string[] rowDelimiters = new string[] { "|||" };
                        string[] columnDelimiters = new string[] { "^^^" };

                        // Получить массив строк из макропеременной
                        String[] rowsXml = macroVarValue.Split(rowDelimiters, StringSplitOptions.None);
                        int i = 0;
                        while (i < rowsXml.Length)
                        {
                            // Получить строку
                            String rowXml = rowsXml[i];

                            // Добавить ноду строки таблицы
                            TableRow newTableRow = table.AppendChild(new TableRow());

                            // Получить из строки массив ячеек
                            String[] cellsXml = rowXml.Split(columnDelimiters, StringSplitOptions.None);

                            int j = 0;
                            while (j < cellsXml.Length)
                            {
                                // Получить ячейку
                                String cellXml = cellsXml[j];

                                // Убрать символ CRLF в конце строки
                                cellXml = cellXml.TrimEnd(new char[] { '\n', '\r' });

                                // Добавить ноду ячейку в строку таблицы
                                TableCell newTableCell = newTableRow.AppendChild(new TableCell());

                                // Вставить текст
                                InsertText(cellXml, paragraphCellStyles[j], newTableCell);

                                j++;
                            }

                            i++;
                        }
                    }

                    /* Картинки */
                    if (sdtProperties.GetFirstChild<SdtContentPicture>() != null)
                    {
                        // Получить путь к файлу
                        String imageFilePath = macroVarValue;

                        // Получить расширение файла
                        String extension = System.IO.Path.GetExtension(imageFilePath).ToLower();
                        ImagePartType imagePartType;
                        switch (extension)
                        {
                            case "jpeg":
                                imagePartType = ImagePartType.Jpeg;
                                break;
                            case "jpg":
                                imagePartType = ImagePartType.Jpeg;
                                break;
                            case "png":
                                imagePartType = ImagePartType.Png;
                                break;
                            case "bmp":
                                imagePartType = ImagePartType.Bmp;
                                break;
                            case "gif":
                                imagePartType = ImagePartType.Gif;
                                break;
                            default:
                                imagePartType = ImagePartType.Jpeg;
                                break;
                        };

                        // Добавить ImagePart в документ
                        ImagePart imagePart = wdDoc.MainDocumentPart.AddImagePart(imagePartType);

                        // Получить картинку
                        using (FileStream stream = new FileStream(imageFilePath, FileMode.Open))
                        {
                            imagePart.FeedData(stream);
                        }

                        // Вычислить width и height
                        Bitmap img = new Bitmap(imageFilePath);
                        var widthPx = img.Width;
                        var heightPx = img.Height;
                        var horzRezDpi = img.HorizontalResolution;
                        var vertRezDpi = img.VerticalResolution;
                        const int emusPerInch = 914400;
                        const int emusPerCm = 360000;
                        var widthEmus = (long)(widthPx / horzRezDpi * emusPerInch);
                        var heightEmus = (long)(heightPx / vertRezDpi * emusPerInch);

                        // Получить ID ImagePart
                        string relationShipId = wdDoc.MainDocumentPart.GetIdOfPart(imagePart);

                        Paragraph paragraph = sdtContentBlock.GetFirstChild<Paragraph>();
                        Run run = paragraph.GetFirstChild<Run>();
                        Drawing drawing = run.GetFirstChild<Drawing>();
                        Inline inline = drawing.GetFirstChild<Inline>();
                        Graphic graphic = inline.GetFirstChild<Graphic>();
                        GraphicData graphicData = graphic.GetFirstChild<GraphicData>();
                        Picture pic = graphicData.GetFirstChild<Picture>();
                        BlipFill blipFill = pic.GetFirstChild<BlipFill>();
                        Blip blip = blipFill.GetFirstChild<Blip>();

                        string prefix = "r";
                        string localName = "embed";
                        string namespaceUri = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                        OpenXmlAttribute oldEmbedAttribute = blip.GetAttribute("embed", namespaceUri);

                        IList<OpenXmlAttribute> attributes = blip.GetAttributes();

                        if (oldEmbedAttribute != null)
                        {
                            attributes.Remove(oldEmbedAttribute);
                        }

                        // Удалить хз что, выявлено практическим путем
                        blipFill.RemoveAllChildren<SourceRectangle>();

                        // Установить новую картинку
                        blip.SetAttribute(new OpenXmlAttribute(prefix, localName, namespaceUri, relationShipId));
                        blip.SetAttribute(new OpenXmlAttribute("cstate", "", "print"));

                        // Подогнать размеры
                        Extent extent = inline.GetFirstChild<Extent>();

                        OpenXmlAttribute oldCxExtent = extent.GetAttribute("cx", "");
                        if (oldCxExtent != null)
                        {
                            var maxWidthEmus = long.Parse(oldCxExtent.Value);
                            if (widthEmus > maxWidthEmus)
                            {
                                var ratio = (heightEmus * 1.0m) / widthEmus;
                                widthEmus = maxWidthEmus;
                                heightEmus = (long)(widthEmus * ratio);
                            }

                            extent.GetAttributes().Remove(oldCxExtent);
                        }

                        OpenXmlAttribute oldCyExtent = extent.GetAttribute("cy", "");
                        if (oldCyExtent != null)
                        {
                            extent.GetAttributes().Remove(oldCyExtent);
                        }

                        extent.SetAttribute(new OpenXmlAttribute("cx", "", widthEmus.ToString()));
                        extent.SetAttribute(new OpenXmlAttribute("cy", "", heightEmus.ToString()));

                        ShapeProperties shapeProperties = pic.GetFirstChild<ShapeProperties>();
                        Transform2D transform2D = shapeProperties.GetFirstChild<Transform2D>();
                        Extents extents = transform2D.GetFirstChild<Extents>();

                        OpenXmlAttribute oldCxExtents = extents.GetAttribute("cx", "");
                        if (oldCxExtents != null)
                        {
                            extents.GetAttributes().Remove(oldCxExtents);
                        }

                        OpenXmlAttribute oldCyExtents = extents.GetAttribute("cy", "");
                        if (oldCyExtents != null)
                        {
                            extents.GetAttributes().Remove(oldCyExtents);
                        }

                        extents.SetAttribute(new OpenXmlAttribute("cx", "", widthEmus.ToString()));
                        extents.SetAttribute(new OpenXmlAttribute("cy", "", heightEmus.ToString()));

                        // Удалить placeholder
                        ShowingPlaceholder showingPlaceholder = sdtProperties.GetFirstChild<ShowingPlaceholder>();
                        if (showingPlaceholder != null)
                        {
                            sdtProperties.RemoveChild<ShowingPlaceholder>(showingPlaceholder);
                        }
                    }

                    /* Повторяющийся раздел */
                    if (sdtProperties.GetFirstChild<SdtRepeatedSection>() != null)
                    {
                        // Представить repeatedSection как новый xml документ (сделать корнем)
                        XmlDocument repeatedSectionXml = new XmlDocument();
                        repeatedSectionXml.LoadXml(macroVarValue);

                        // Получить корневой элемент repeatedSection
                        XmlElement rootRepeatedSectionXml = repeatedSectionXml.DocumentElement;

                        // Получить количество repeatedSectionItem
                        XmlNodeList repeatedSectionItems = rootRepeatedSectionXml.SelectNodes("repeatedSectionItem");
                        int repeatedItemCount = repeatedSectionItems.Count;

                        Console.WriteLine("Количество repeatedSectionItem: " + repeatedItemCount);

                        /* Блок клонирования ноды повтор. раздела до нужного количества */
                        for (int  i = 0; i < repeatedItemCount; i++)
                        {
                            XmlElement macroListRepeatedSectionItem = rootRepeatedSectionXml.SelectSingleNode(String.Format(@"repeatedSectionItem[@id=""{0}""]", i)) as XmlElement;
                            Console.WriteLine("Item " + i + ": " + macroListRepeatedSectionItem.OuterXml);

                            SdtContentBlock sdtContentBlockRepeatedSectionItem = sdtContentBlock.Elements<SdtBlock>().Last<SdtBlock>().GetFirstChild<SdtContentBlock>();

                            Fill(sdtContentBlockRepeatedSectionItem, macroListRepeatedSectionItem, wdDoc);

                            if (i + 1 < repeatedItemCount)
                            {
                                SdtBlock clonedRepeatedSectionItem = sdtContentBlock.GetFirstChild<SdtBlock>().Clone() as SdtBlock;
                                sdtContentBlock.AppendChild<SdtBlock>(clonedRepeatedSectionItem);
                            }
                        }
                        /**/

                        //Fill(sdtContentBlock, macroListRepeatedSection, wdDoc);
                    }
                }

                Console.WriteLine();
            }
        }

        public static String FindMacroVar(XmlElement macroListXml, String name, String type = "")
        {
            // Найти в macroListXml Node макропеременной
            // Поиск по тегу и имени
            //XmlNode macroVarNode = macroListXml.SelectSingleNode(String.Format(@"richText[@name=""{0}""]", tag.Val));
            // Поиск только по имени
            XmlNode macroVarNode = macroListXml.SelectSingleNode(String.Format(@"*[@name=""{0}""]", name));

            String macroVarValue = "";
            // Если макропеременная задана
            if (macroVarNode != null)
            {
                // Получить ноду CDATA со значением
                XmlNode firstChild = macroVarNode.FirstChild;

                if (firstChild.NodeType == XmlNodeType.CDATA)
                {
                    // Получить значение макропеременной из CDATA
                    macroVarValue = System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(firstChild.Value));
                }
                else
                {
                    macroVarValue = macroVarNode.OuterXml;
                }

                return macroVarValue;
            }
            else
            {
                return null;
            }
        }

        public static void InsertText(String text, String styleXml, OpenXmlElement parentNode)
        {
            // Разделить строку на подстроки 
            //String[] runs = macroVarValue.Split(new char[] { '\r', '\n' });
            //String[] runs = Regex.Split(macroVarValue, "\r\n|\r|\n");
            String[] runs = Regex.Split(text, "\r\n");

            // Вставить текст
            foreach (String run in runs)
            {
                // Создать новый Paragraph
                Paragraph newParagraph = parentNode.AppendChild(new Paragraph());

                // Вернуть старый стиль в Paragraph
                ParagraphProperties newParagraphProperties = newParagraph.AppendChild(new ParagraphProperties());
                newParagraphProperties.InnerXml = styleXml;

                // Создать новый Run
                Run newRun = newParagraph.AppendChild(new Run());

                // Вернуть старый стиль в Run
                RunProperties newRunProperties = newRun.AppendChild(new RunProperties());
                newRunProperties.InnerXml = styleXml;

                // Вставить в него текст
                newRun.AppendChild(new Text(run));
            }
        }
    }
}
