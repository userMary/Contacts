using Microsoft.AspNetCore.Mvc;
using Contacts.Models;
using System.Text;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using iTextSharp.text.pdf;

using DocumentFormat.OpenXml;
using Paragraph = iTextSharp.text.Paragraph;
using OfficeOpenXml;

namespace Contacts.Controllers
{
    public class ContactsController : Controller
    {
        private static List<Contact> _contacts = new List<Contact>
        {
            new Contact { Name = "Иван Иванов", MobilePhone = "+7 (900) 123-45-67", AlternateMobilePhone = "+7 (900) 765-43-21", Email = "ivan@example.com", Description = "Друг из школы" },
            new Contact { Name = "Мария Петрова", MobilePhone = "+7 (900) 234-56-78", AlternateMobilePhone = "+7 (900) 841-75-24", Email = "maria@example.com", Description = "Коллега" },
            new Contact 
                    {
                        Name = "Anna Petrova",
                        MobilePhone = "+7 (915) 123-4567",
                        AlternateMobilePhone = "+7 (495) 765-4321",
                        Email = "anna.petrova@example.com",
                        Description = "Коллега из отдела маркетинга"
                    },
            new Contact
                    {
                        Name = "Ivan Ivanov",
                        MobilePhone = "+7 (926) 234-5678",
                        AlternateMobilePhone = "+7 (495) 876-5432",
                        Email = "ivan.ivanov@example.com",
                        Description = "Бизнес-партнер"
                    },
            new Contact
                    {
                        Name = "Elena Smirnova",
                        MobilePhone = "+7 (916) 345-6789",
                        AlternateMobilePhone = "+7 (495) 678-9123",
                        Email = "elena.smirnova@example.com",
                        Description = "Подруга из университета"
                    },
            new Contact
                    {
                        Name = "Sergey Kuznetsov",
                        MobilePhone = "+7 (925) 456-7890",
                        AlternateMobilePhone = "+7 (495) 543-2109",
                        Email = "sergey.kuznetsov@example.com",
                        Description = "Партнер по проекту"
                    },
            new Contact
                    {
                        Name = "Olga Sokolova",
                        MobilePhone = "+7 (927) 567-8901",
                        AlternateMobilePhone = "+7 (495) 432-1098",
                        Email = "olga.sokolova@example.com",
                        Description = "HR менеджер"
                    },
            new Contact
                    {
                        Name = "Dmitry Morozov",
                        MobilePhone = "+7 (919) 678-9012",
                        AlternateMobilePhone = "+7 (495) 321-0987",
                        Email = "dmitry.morozov@example.com",
                        Description = "Клиент"
                    },
            new Contact
                    {
                        Name = "Nadezhda Lebedeva",
                        MobilePhone = "+7 (918) 789-0123",
                        AlternateMobilePhone = "+7 (495) 210-9876",
                        Email = "nadezhda.lebedeva@example.com",
                        Description = "Руководитель отдела"
                    },
            new Contact
                    {
                        Name = "Viktor Popov",
                        MobilePhone = "+7 (921) 890-1234",
                        AlternateMobilePhone = "+7 (495) 109-8765",
                        Email = "viktor.popov@example.com",
                        Description = "Финансовый консультант"
                    },
            new Contact
                    {
                        Name = "Natalya Volchkova",
                        MobilePhone = "+7 (924) 901-2345",
                        AlternateMobilePhone = "+7 (495) 908-7654",
                        Email = "natalya.volchkova@example.com",
                        Description = "Старый друг"
                    },
            new Contact
                    {
                        Name = "Mikhail Romanov",
                        MobilePhone = "+7 (923) 012-3456",
                        AlternateMobilePhone = "+7 (495) 807-6543",
                        Email = "mikhail.romanov@example.com",
                        Description = "Коллега из IT отдела"
                    },
            new Contact
                    {
                        Name = "Tatiana Zueva",
                        MobilePhone = "+7 (922) 123-4567",
                        AlternateMobilePhone = "+7 (495) 706-5432",
                        Email = "tatiana.zueva@example.com",
                        Description = "Соседка"
                    },
            new Contact
                    {
                        Name = "Andrey Karpov",
                        MobilePhone = "+7 (929) 234-5678",
                        AlternateMobilePhone = "+7 (495) 605-4321",
                        Email = "andrey.karpov@example.com",
                        Description = "Технический директор"
                    },
            new Contact
                    {
                        Name = "Ekaterina Guseva",
                        MobilePhone = "+7 (928) 345-6789",
                        AlternateMobilePhone = "+7 (495) 504-3210",
                        Email = "ekaterina.guseva@example.com",
                        Description = "Знакомая с тренировки"
                    },
            new Contact
                    {
                        Name = "Alexey Fedorov",
                        MobilePhone = "+7 (917) 456-7890",
                        AlternateMobilePhone = "+7 (495) 403-2109",
                        Email = "alexey.fedorov@example.com",
                        Description = "Юрист компании"
                    },
            new Contact
                    {
                        Name = "Svetlana Ivanova",
                        MobilePhone = "+7 (920) 567-8901",
                        AlternateMobilePhone = "+7 (495) 302-1098",
                        Email = "svetlana.ivanova@example.com",
                        Description = "Куратор курса"
                    }
        };

        public IActionResult Index()
        {
            return View(_contacts);
        }

        [HttpPost]
        public IActionResult SaveContacts(string format)
        {
            if (format == "txt") return SaveAsText();
            //if (format == "docx") return SaveAsDocx();
            if (format == "pdf") return SaveAsPdf();
            if (format == "xlsx") return SaveAsXlsx();
            return BadRequest("Неподдерживаемый формат");
        }

        private FileResult SaveAsText()
        {
            var builder = new StringBuilder();
            foreach (var contact in _contacts)
            {
                builder.AppendLine($"{contact.Name}|{contact.MobilePhone}|{contact.AlternateMobilePhone}|" +
                    $"{contact.Email}|{contact.Description}");
                //builder.AppendLine($"Mobile Phone: {contact.MobilePhone}");
                //builder.AppendLine($"Alternate Mobile Phone: {contact.AlternateMobilePhone}");
                //builder.AppendLine($"Email: {contact.Email}");
                //builder.AppendLine($"Description: {contact.Description}");
                builder.AppendLine();
            }

            var fileBytes = Encoding.UTF8.GetBytes(builder.ToString());
            return File(fileBytes, "text/plain", "contacts_data.txt");
        }

        //private FileResult SaveAsDocx()
        //{
        //    using var stream = new MemoryStream();
        //    using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

        //    // Добавляем основную часть документа
        //    var mainPart = doc.AddMainDocumentPart();
        //    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
        //    var body = mainPart.Document.AppendChild(new Body());

        //    // Создаем содержимое документа
        //    foreach (var contact in _contacts)
        //    {
        //        body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text($"Name: {contact.Name}"))));
        //        body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text($"Mobile Phone: {contact.MobilePhone}"))));
        //        body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text($"Alternate Mobile Phone: {contact.AlternateMobilePhone}"))));
        //        body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text($"Email: {contact.Email}"))));
        //        body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text($"Description: {contact.Description}"))));
        //        body.AppendChild(new DocumentFormat.OpenXml.Drawing.Paragraph(new Run(new Text(" ")))); // Пустая строка для разделения контактов
        //    }

        //    // Сохраняем изменения в основной части
        //    mainPart.Document.Save();

        //    // Перемещаем поток на начало
        //    stream.Seek(0, SeekOrigin.Begin);

        //    // Возвращаем файл
        //    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "contacts_data.docx");
        //}

        private FileResult SaveAsPdf()
        {
            using var stream = new MemoryStream();
            var document = new iTextSharp.text.Document();

            // Настройка PdfWriter
            PdfWriter writer = PdfWriter.GetInstance(document, stream);
            writer.CloseStream = false;

            document.Open(); // Открываем документ для записи

            // Добавляем содержимое документа PDF
            foreach (var contact in _contacts)
            {
                document.Add(new Paragraph($"Name: {contact.Name}"));
                document.Add(new Paragraph($"Mobile Phone: {contact.MobilePhone}"));
                document.Add(new Paragraph($"Alternate Mobile Phone: {contact.AlternateMobilePhone}"));
                document.Add(new Paragraph($"Email: {contact.Email}"));
                document.Add(new Paragraph($"Description: {contact.Description}"));
                document.Add(new Paragraph(" ")); // Пустая строка для разделения контактов
            }

            document.Close(); // Закрываем документ

            return File(stream.ToArray(), "application/pdf", "contacts_data.pdf");
        }

        private FileResult SaveAsXlsx()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Contacts");

            worksheet.Cell(1, 1).Value = "Name";
            worksheet.Cell(1, 2).Value = "Mobile Phone";
            worksheet.Cell(1, 3).Value = "Alternate Mobile Phone";
            worksheet.Cell(1, 4).Value = "Email";
            worksheet.Cell(1, 5).Value = "Description";

            for (int i = 0; i < _contacts.Count; i++)
            {
                worksheet.Cell(i + 2, 1).Value = _contacts[i].Name;
                worksheet.Cell(i + 2, 2).Value = _contacts[i].MobilePhone;
                worksheet.Cell(i + 2, 3).Value = _contacts[i].AlternateMobilePhone;
                worksheet.Cell(i + 2, 4).Value = _contacts[i].Email;
                worksheet.Cell(i + 2, 5).Value = _contacts[i].Description;
            }

            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "contacts_data.xlsx");
        }


        //[HttpPost]
        //public IActionResult LoadFromPdf(IFormFile file)
        //{
        //    if (file == null || file.Length == 0)
        //        return RedirectToAction("Index");

        //    var contacts = new List<Contact>();

        //    using (var reader = new PdfReader(file.OpenReadStream()))
        //    {
        //        for (int i = 1; i <= reader.NumberOfPages; i++)
        //        {
        //            var pageText = PdfTextExtractor.GetTextFromPage(reader, i);
        //            var lines = pageText.Split('\n');
        //            foreach (var line in lines)
        //            {
        //                var fields = line.Split(':');
        //                if (fields.Length == 5)
        //                {
        //                    contacts.Add(new Contact
        //                    {
        //                        Name = fields[0].Trim(),
        //                        MobilePhone = fields[1].Trim(),
        //                        AlternateMobilePhone = fields[2].Trim(),
        //                        Email = fields[3].Trim(),
        //                        Description = fields[4].Trim()
        //                    });
        //                }
        //            }
        //        }
        //    }

        //    _contacts = contacts;
        //    return RedirectToAction("Index");
        //}

        [HttpPost]
        public IActionResult LoadFromXlsx(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return RedirectToAction("Index");

            var contacts = new List<Contact>();

            using (var stream = file.OpenReadStream())
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets[0];
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    contacts.Add(new Contact
                    {
                        Name = worksheet.Cells[row, 1].Text,
                        MobilePhone = worksheet.Cells[row, 2].Text,
                        AlternateMobilePhone = worksheet.Cells[row, 3].Text,
                        Email = worksheet.Cells[row, 4].Text,
                        Description = worksheet.Cells[row, 5].Text
                    });
                }
            }

            _contacts = contacts;
            return RedirectToAction("Index");
        }

        [HttpPost]
        public IActionResult LoadFromTxt(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return RedirectToAction("Index");

            var contacts = new List<Contact>();

            using (var reader = new StreamReader(file.OpenReadStream()))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var fields = line.Split('|');
                    if (fields.Length == 5)
                    {
                        contacts.Add(new Contact
                        {
                            Name = fields[0].Trim(),
                            MobilePhone = fields[1].Trim(),
                            AlternateMobilePhone = fields[2].Trim(),
                            Email = fields[3].Trim(),
                            Description = fields[4].Trim()
                        });
                    }
                }
            }

            _contacts = contacts;
            return RedirectToAction("Index");
        }



    }
}
