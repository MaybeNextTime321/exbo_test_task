using System;
using System.Text.Json;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
class Program
{
    static void Main(string[] args)
    {
        /*
            Словари для хранения данных
            Items хранит данные с items.csv
            Rewards хранит данные с file.json и после включает в себя данные из items.csv
            FirstTaskResult хранит данные для первой задачи
            SecondTaskUnlistedReward содержит данные, которые не находятся ни в одном листе

         */
        Dictionary<string, RewardDetails> Items = new Dictionary<string, RewardDetails>(); ;
        Dictionary<string, Reward> Rewards = new Dictionary<string, Reward>(); ;
        Dictionary<string, Reward> FirstTaskResult = new Dictionary<string, Reward>(); ;
        Dictionary<string, Reward> SecondTaskUnlistedReward = new Dictionary<string, Reward>(); ;


        /*
            resultDirectory хранит поддиректорию, в которую будет записан результат
            Чтение .json файлов и инициализация пути для .csv и .xlsx
            itemsPath хранит путь к файлу с данными айтемов
            firstTaskResultFileName название файла, в которое будет выведен результат первого задания
            secondTaskResultFileName название файла, в которое будет выведен результат второго задания
            fileJson и taskJson читает представленные файл
         */
        string resultDirectory = "Result";
        string itemsPath = "items.csv";
        string firstTaskResultFileName = $"{resultDirectory}/firstTask.json";
        string secondTaskResultFileName = $"{resultDirectory}/secondTask.xlsx";
        string fileJson = File.ReadAllText("file.json");
        string taskJson = File.ReadAllText("task.json");

        // Относительный путь, который мы проверяем
        string relativeResultPath = @$"{resultDirectory}\";

        // Запись в переменную полного пути
        string fullResultPath = Path.Combine(Directory.GetCurrentDirectory(), relativeResultPath);

        //Если такой директории нет, создать ее
        if (!Directory.Exists(fullResultPath))
        {
            Directory.CreateDirectory(fullResultPath);
            Console.WriteLine($"Создана поддиректория Result для сохранения результатов") ;
        }

        Console.WriteLine($"Начинаю анализ файла {itemsPath}");

        // Открываем файл items.csv для чтения
        using (var reader = new StreamReader(itemsPath))
        {
            // Инициализация переменной для хранения строки и запись в нее данных, пока в таблице есть данные
            string line = "";
            while ((line = reader.ReadLine()) != null)
            {
                // Разделение целой строки на отдельные данные
                string id = line.Split(',')[0];
                int money = Int32.Parse(line.Split(',')[1]);
                int details = Int32.Parse(line.Split(',')[2]);
                int reputation = Int32.Parse(line.Split(',')[3]);

                // Запись полученных данных в RewardDetails
                Items.Add(id, new RewardDetails(money, details, reputation));

            }
        }

        Console.WriteLine($"Анализ файла {itemsPath} успешно завершен");


        
        Console.WriteLine($"Начинаю анализ файла fileJson");

        // Парсинг file.json и получение его корневого элемента
        using JsonDocument doc = JsonDocument.Parse(fileJson);
        JsonElement root = doc.RootElement;

        // В корневом элементе мы проходимся по каждому отдельному элементу
        foreach(JsonProperty property in root.EnumerateObject())
        {

            // Чтение необходимых данных в каждом элементе
            string objectName = property.Name;
            int id = property.Value.GetProperty("id").GetInt32();
            string reward = property.Value.GetProperty("reward").GetString();
            int weight = property.Value.GetProperty("weight").GetInt32();

            // Создание новой записи и запись ее в Rewards и SecondTaskUnlistedReward
            // (мы сначала добавляем туда все записи, а потом убираем те, которые есть хоть в одном list)
            Reward newItem = new Reward(id, weight, reward);
            Rewards.Add(objectName, newItem);
            SecondTaskUnlistedReward.Add(objectName, newItem);


            // Если в items.csv были данные для этого reward, то записываем их
            if(Items.ContainsKey(reward))
            {
                newItem.SetRewardDetails(Items[reward]);
            }

        }

        Console.WriteLine($"Анализ файла fileJson успешно завершен");

        // Разрешение лицензии на использование EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Создание файла Excel
        var package = new ExcelPackage();

        // Создаем новый лист и заголовки
        var worksheet = package.Workbook.Worksheets.Add("SecondTask");
        worksheet.Cells[1,1].Value = "list_name";
        worksheet.Cells[1,2].Value = "object_name";
        worksheet.Cells[1,3].Value = "reward_key";
        worksheet.Cells[1,4].Value = "money";
        worksheet.Cells[1,5].Value = "details";
        worksheet.Cells[1,6].Value = "reputation";
        worksheet.Cells[1,7].Value = "isUsed";
        
        //Переменная, ответственная за текущую строку
        int row = 2;

        Console.WriteLine($"Начинаю анализ файла taskJson");

        // Парсинг документа с тасками
        using JsonDocument docTask = JsonDocument.Parse(taskJson);
        JsonElement rootTask = docTask.RootElement;

        // Парсинг проходимся по всем его элементам
        foreach (JsonProperty property in rootTask.EnumerateObject())
        {
            // В каждом элементе достаем list и проходимся по его значениям
            JsonElement listElement = property.Value.GetProperty("list");
            foreach(JsonElement item in listElement.EnumerateArray())
            {
                // Получаем данные этой награды
                Reward data = Rewards[item.ToString()];

                // Записываем в таблицу list_name, object_name и isUsed
                worksheet.Cells[row, 1].Value = property.Name;
                worksheet.Cells[row, 2].Value = item.ToString();
                worksheet.Cells[row, 7].Value = 1;
                
                // Если в результатах для 1 задания нет этой награды, то добавляем ее 
                // А также помечаем, что она была использована хоть раз (для isUsed = 0)
                if(!FirstTaskResult.ContainsKey(item.ToString()))
                {
                    if(data != null)
                    {
                        FirstTaskResult.Add(item.ToString(), data);
                        SecondTaskUnlistedReward.Remove(item.ToString());
                    }
                }

                // Если мы получили данные награды, то записываем в таблицу reward, money, details, reputation
                if (data!=null )
                {
                    worksheet.Cells[row, 3].Value = data.reward;
                    worksheet.Cells[row, 4].Value = data.money;
                    worksheet.Cells[row, 5].Value = data.details;
                    worksheet.Cells[row, 6].Value = data.reputation;
                }

                // Не забываем про перенос на следующую строку
                row++;
            }

        }

        Console.WriteLine($"Анализ файла taskJson успешно завершен");
        Console.WriteLine($"Добавляю в таблицу неиспользованные награды");

        // Проходимся по всем неиспользованным в list наградах
        foreach (KeyValuePair<string,Reward> reward in SecondTaskUnlistedReward)
        {
            // На всякий случай пропускаем те, что использовались, но по факту тут всегда false
            // Однако много проверок не бывает
            if(FirstTaskResult.ContainsKey(reward.Key)) 
            {
                continue;
            }

            //Если мы получили данные по этому ключу, то записываем их в таблицу
            if(reward.Value != null)
            { 
                worksheet.Cells[row, 2].Value = reward.Key;
                worksheet.Cells[row, 3].Value = reward.Value.reward;
                worksheet.Cells[row, 4].Value = reward.Value.money;
                worksheet.Cells[row, 5].Value = reward.Value.details;
                worksheet.Cells[row, 6].Value = reward.Value.reputation;
                worksheet.Cells[row, 7].Value = 0;
            }
            
            // Не забываем перенести номер строки
            row++;
        }
        Console.WriteLine($"Неиспользованные награды успешно добавлены в таблицу");

        Console.WriteLine($"Создаю итоговый файл 1 задания {firstTaskResultFileName}");
        // Серелизуем данные, которые сохранили для первой задачи и сохраняем их в файле
        string firstTaskJSON = JsonSerializer.Serialize(FirstTaskResult, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(firstTaskResultFileName, firstTaskJSON);
        Console.WriteLine($"Итоговый файл 1 задания {firstTaskResultFileName} успешно создан");


        Console.WriteLine($"Сортирую таблицу по list_name");
        // Сортируем данные в таблице второго задания по 1 колонке (list_name)
        worksheet.Cells[$"A2:G{row}"].Sort(x => x.SortBy.Column(0));
        Console.WriteLine($"Таблица успешно отсортирована по list_name");

        Console.WriteLine($"Сохраняю таблицу {secondTaskResultFileName}");
        // Сохраняем таблицу для 2 задачи
        File.WriteAllBytes(secondTaskResultFileName, package.GetAsByteArray());
     
        Console.WriteLine($"Файл Excel создан: {secondTaskResultFileName}");

    }
}

//Класс с данными о награду
class Reward
{
    // Конструктор с данными, которые мы получаем из file.json
    public Reward(int id, int weight, string reward)
    {
        _id = id;
        _weight = weight;
        _reward = reward;
    }
    
    // Запись данных, которые мы получаем из items.csv
    public void SetRewardDetails(RewardDetails details)
    {
        _rewardDetails = details;
    }

    // Установка публичных полей для серелизации в JOSN
    public string reward => _reward;
    public int money => _rewardDetails.GetMoney();
    public int details => _rewardDetails.GetDetails();
    public int reputation => _rewardDetails.GetRetutation();

    // Приватные поля 

    private int _id;
    private int _weight;
    private string _reward = "";
    private RewardDetails _rewardDetails;
}

// Структура с данными из items.csv
struct RewardDetails
{
    // Публичные гетеры для получены данных в Rewards
    public int GetMoney()
    {
        return _money;
    }

    public int GetDetails()
    {
        return _details;
    }

    public int GetRetutation()
    {
        return _reputation;
    }

    // Конструктор с полями, которые мы получаем из items.csv
    public RewardDetails(int money, int details, int reputation)
    {
        _money = money;
        _details = details;
        _reputation = reputation;
    }

   // Приватные поля структуры
   private int _money;
   private int _details;
   private int _reputation;
}
