using TableScraping;
using System.Runtime.Serialization.Json;
using System.Reflection;
using System.Collections;
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

//string pathCSV = @"C:\Users\79996\Downloads\import_ou_csv.csv";
string pathCSV = @"C:\Users\79996\Downloads\table.csv";
char delimiter = ';';
//CSVParse parse1 = new CSVParse(pathCSV);
//var a = parse1.Scraping();
//parse1.SerializeIntoJSON();  // сериализация

//
//ПАРСИНГ CSV
//


var parse = new CSVParse();
var method = typeof(CSVParse).GetMethod("Scraping");
var data = method.Invoke(parse, new object[] { pathCSV });
IEnumerable enumerable = data as IEnumerable;
if (enumerable != null)
{

    string[] array = enumerable as string[];
    for (int i = 1; i < array.Length; i++)
    {
        string[] raw = array[i].Split(delimiter);
        CSV model = new CSV
        {
            Name = raw[1],
            Id = Convert.ToInt32(raw[0])
        };

        Console.WriteLine(model.Id + " " + model.Name);
    }
}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


string pathXLSX = @"D:\79996\Documents\хз\Лист Microsoft Excel.xlsx";
//XLSXParse parse = new XLSXParse(pathXLSX);
//var b = parse.Scraping();
//parse.SerializeIntoJSON(); //сериализация

var parseExcel = new XLSXParse();
var methodExcel = typeof(XLSXParse).GetMethod("Scraping");
var dataExcel = methodExcel.Invoke(parseExcel, new object[] { pathXLSX });
IEnumerable enumerableExcel = dataExcel as IEnumerable;
if (enumerableExcel != null)
{
    string[] array = enumerableExcel as string[];
    if (array != null) 
    {
        for(int i = 1; i < array.Length; i++)
        {
            var arrayss = array[i].Split(delimiter, StringSplitOptions.RemoveEmptyEntries);
            for (int j = 0; j < arrayss.Length; j++)
            {
                XLSX xlsx = new XLSX()
                {
                    NameAndFormOfEvent = arrayss[0],
                    OnlineOrOffline = arrayss[1],
                    ReaderAssignment = arrayss[2],
                    ShortDescription = arrayss[3],
                };
                //Console.WriteLine(xlsx.NameAndFormOfEvent);
                //var tb_FullName = string.Join(" ", new string[] { xlsx.NameAndFormOfEvent, xlsx.OnlineOrOffline, xlsx.ReaderAssignment,xlsx.ShortDescription }.Where(s => s?.Length > 0));
                //Console.WriteLine(tb_FullName);
                Console.WriteLine(xlsx.NameAndFormOfEvent +" "+ xlsx.OnlineOrOffline + " " + xlsx.ReaderAssignment + " " + xlsx.ShortDescription);
            }
        }
    }
}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
class CSV
{
    [TableModel("ID", "int")]
    public int Id { get; set; }
    [TableModel("Name", "string")]
    public string? Name { get; set; }
}

class XLSX
{
    [TableModel("Название и форма мероприятия", "string")]
    public string? NameAndFormOfEvent{ get; set; }
    [TableModel("Онлайн/оффлайн", "string")]
    public string? OnlineOrOffline { get; set; }
    [TableModel("Читатель, назначение", "string")]
    public string? ReaderAssignment { get; set; }
    [TableModel("Краткое описание", "string")]
    public string? ShortDescription { get; set; }
}