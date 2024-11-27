
using ExcelToMailSign.Models;
using OfficeOpenXml;
using System.Text.RegularExpressions;

class Program
{

    // Read Excell
    // Replace HTML with readed Excel
    // Send mail replaced HTMLs

    static void Main(string[] args)
    {

      var employess = ReadEmployeeExcel();
      ReplaceHTML(employess);

    }

    public static List<Employee> ReadEmployeeExcel()
    {
        // buraya select FileDialog eklenebilir. Zaman kaybetmemek için bakmadım
        string filePath = @"C:\Users\PC\Desktop\Projects\ExcelToMailSign\SampleContacts.xlsx";

        if (!File.Exists(filePath))
        {
            Console.WriteLine("Dosya bulunamadı. Lütfen doğru bir yol girin.");
            return null;
        }

        List<Employee> employees = new List<Employee>();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                var employee = new Employee
                {
                    FullName = worksheet.Cells[row, 1].Text,
                    Email = worksheet.Cells[row, 2].Text,
                    Title = worksheet.Cells[row, 3].Text,
                    Phone = worksheet.Cells[row, 4].Text
                };
                employees.Add(employee);
            }
        }

        //foreach (var employee in employees)
        //{

        //    Console.WriteLine($"Name: {employee.FullName}, Email: {employee.Email}, Title: {employee.Title}, Phone: {employee.Phone}");
        //}

        return employees;
    }

    public static bool ReplaceHTML(List<Employee> employees)
    {
        string htmlTemplate = File.ReadAllText(@"C:/Users/PC/Desktop/Projects/ExcelToMailSign/HTMLTemplate.html");

        foreach (var employee in employees)
        {
            string personalizedHtml = htmlTemplate;

            var properties = new Dictionary<string, string>
        {
            { "##fullName##", employee.FullName },
            { "##title##", employee.Title },
            { "##email##", employee.Email },
            { "##phoneNumber##", employee.Phone }
        };

            foreach (var placeholder in properties)
            {
                personalizedHtml = personalizedHtml.Replace(placeholder.Key, placeholder.Value);
            }

            string fileName = employee.Email.Split('@')[0];
            fileName = fileName.Replace(" ", "_").Replace(".", "_");

            //kaydetme
            File.WriteAllText($@"C:\Users\PC\Desktop\Projects\ExcelToMailSign\{fileName}.html", personalizedHtml);

            Console.WriteLine($"HTML dosyası oluşturuldu: {fileName}.html");
        }

        return true;
    }

}
