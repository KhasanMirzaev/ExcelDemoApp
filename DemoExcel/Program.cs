using System;
using System.ComponentModel;
using ExcelDemo;
using OfficeOpenXml;

ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
var excelFile = new FileInfo(@"C:\DirectoryExcelFile\DemoExcelFile.xlsx");

var peopleList = GetSetupData();

await SaveExcelFile(peopleList, excelFile);

static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
{
    DeleteIfExists(file);

    using var package = new ExcelPackage(file); //tushunimcha filening open/closelariga javob beradi(excel file yaratdi)

    var workSheet = package.Workbook.Worksheets.Add("Mainreport");//mainreport nomi yangi sheet och(worksheet yaradi)

    var range = workSheet.Cells["A1"].LoadFromCollection(people, true);

    await package.SaveAsync();
}

//berilgan fayl mavjud bo'lsa o'chirib yuboradi
static void DeleteIfExists(FileInfo file)
{
    if (file.Exists)
        file.Delete();
}

//chaqirilganda list qaytaradi(personmodel saqlovchi list)
static List<PersonModel> GetSetupData()
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
