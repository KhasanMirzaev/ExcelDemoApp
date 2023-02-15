using OfficeOpenXml;

namespace ExcelDemo
{
    static class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excelFile = new FileInfo(@"..\..\..\..\..\DirectoryExcelFile\DemoExcelFile.xlsx");

            var peopleList = GetSetupData();

            await SaveExcelFile(peopleList, excelFile);

            List<PersonModel> peopleFromExcel = await LoadExcelFile(excelFile);

            foreach(var person in peopleFromExcel)
            {
                Console.WriteLine($"{person.Id} {person.FirstName} {person.LastName}");
            }
        }

        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            //qaysi worksheet b-n ishlashni tanlaymiz(bu yerda 0-w.sh. tanlandi
            var workSheet = package.Workbook.Worksheets[0];

            int row = 2, col = 1;
            List<PersonModel> people = new List<PersonModel>();

            //tekshirilayotgan celldagi qiymat null yoki bo'sh bo'lganda sikl to'taydi(ya'ni ma'lumot tuagaganda)
            while (string.IsNullOrEmpty(workSheet.Cells[row, col].Value?.ToString()) == false)
            {
                var person = new PersonModel();

                person.Id = int.Parse(workSheet.Cells[row, col].Value.ToString());
                person.FirstName = workSheet.Cells[row, col+1].Value.ToString();
                person.LastName = workSheet.Cells[row, col+2].Value.ToString();

                people.Add(person);
                row++;
            }

            return people;
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            //filening open/closelariga javob beradi(excel file yaratdi)
            using var package = new ExcelPackage(file);

            //mainreport nomi yangi sheet och(worksheet yaradi)
            var workSheet = package.Workbook.Worksheets.Add("Mainreport");

            var range = workSheet.Cells["A1"].LoadFromCollection(people, true);

            await package.SaveAsync();
        }

        //berilgan fayl mavjud bo'lsa o'chirib yuboradi
        private static void DeleteIfExists(FileInfo file)
        {
            if(file.Exists)
                file.Delete();
        }

        //chaqirilganda list qaytaradi(personmodel saqlovchi list)
        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> excelData = new()
            {
                new() {Id = 1, FirstName = "Sharif", LastName="Haqqoniy"},
                new() {Id = 2, FirstName = "Jaloliddin", LastName="Zakiy"},
                new() {Id = 3, FirstName = "Diyor", LastName="Muttaqiy"},
                new() {Id = 4, FirstName = "Jamoliddin", LastName="Horun"},
                new() {Id = 5, FirstName = "Javohir", LastName="Muborak"},
            };

            return excelData;
                
        }
    }    
}
