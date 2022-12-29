using Microsoft.VisualBasic;
using System;
using System.Data.SQLite;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.ReplyMarkups;
using System.Text;
using System.ComponentModel;
using System.Data;


namespace telegram_bot
{
    class Program
    {
        public static SQLiteConnection DB;
        static void Main(string[] args)
        {
            var botClient = new TelegramBotClient("5848911099:AAFoFFWWMR4ZVj9DtUUYHihiziuTtC4JPVI");

            Console.WriteLine($"Бот {botClient.GetMeAsync().Result.FirstName} Почав роботу.");
            var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;
            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = { }
            };
            botClient.StartReceiving(
                HandleUpdatesAsync,
                HandleErrorAsync,
                receiverOptions,
                cancellationToken);
            Console.ReadLine();
            cts.Cancel();
        }
        static async Task HandleUpdatesAsync(ITelegramBotClient botClient, Update update, CancellationToken arg3)
        {
            if (update.Type == UpdateType.Message && update?.Message?.Text != null)
            {
                await HandleMessage(botClient, update.Message, update, arg3);
                return;
            }
        }
        static async Task HandleMessage(ITelegramBotClient botClient, Message message, Update update, CancellationToken arg3)
        {
            if (message.Text == "/start")
            {
                Audit(botClient, update, arg3);
                ReplyKeyboardMarkup keyboard = new(new[]
            {
            new KeyboardButton[] { "А", "Б", "Е" },
            new KeyboardButton[] { "М", "МД", "МК" },
            new KeyboardButton[] { "П", "Т", "Ф"  }
        })
                {
                    ResizeKeyboard = true
                };
                await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть спеціальність", replyMarkup: keyboard);
                return;
            }
            switch (message.Text)
            {
                case "А":
                    await ChangeSpeciality(botClient, update, "А");

                    ReplyKeyboardMarkup keyboard_a = new(new[]
                {
                new KeyboardButton[] { "11", "12" },
                new KeyboardButton[] { "21", "22" },
                new KeyboardButton[] { "31", "32" },
                new KeyboardButton[] { "41" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_a);
                    return;

                case "Б":
                    await ChangeSpeciality(botClient, update, "Б");

                    ReplyKeyboardMarkup keyboard_b = new(new[]
                {
                new KeyboardButton[] { "11", "21", "31" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_b);
                    return;

                case "Е":
                    await ChangeSpeciality(botClient, update, "Е");

                    ReplyKeyboardMarkup keyboard_e = new(new[]
                {
                new KeyboardButton[] { "11", "21" },
                new KeyboardButton[] { "31", "41" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_e);
                    return;

                case "М":
                    await ChangeSpeciality(botClient, update, "М");

                    ReplyKeyboardMarkup keyboard_m = new(new[]
                {
                new KeyboardButton[] { "11", "21" },
                new KeyboardButton[] { "31", "41" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_m);
                    return;

                case "МД":
                    await ChangeSpeciality(botClient, update, "МД");

                    ReplyKeyboardMarkup keyboard_md = new(new[]
                {
                new KeyboardButton[] { "11", "21", "31" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_md);
                    return;

                case "МК":
                    await ChangeSpeciality(botClient, update, "МК");

                    ReplyKeyboardMarkup keyboard_mk = new(new[]
                {
                 new KeyboardButton[] { "11", "21" },
                 new KeyboardButton[] { "31", "41" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_mk);
                    return;

                case "П":
                    await ChangeSpeciality(botClient, update, "П");

                    ReplyKeyboardMarkup keyboard_p = new(new[]
                {
                new KeyboardButton[] { "11", "12", "13" },
                new KeyboardButton[] { "21", "22", "23" },
                new KeyboardButton[] { "31", "32" },
                new KeyboardButton[] { "41", "42" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_p);
                    return;

                case "Т":
                    await ChangeSpeciality(botClient, update, "Т");

                    ReplyKeyboardMarkup keyboard_t = new(new[]
                {
                new KeyboardButton[] { "21", "31", "41" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_t);
                    return;

                case "Ф":
                    Console.WriteLine("Ф");
                    await ChangeSpeciality(botClient, update, "Ф");

                    ReplyKeyboardMarkup keyboard_f = new(new[]
                {
                new KeyboardButton[] { "11", "21", "31" }
             })
                    {
                        ResizeKeyboard = true
                    };
                    await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть групу", replyMarkup: keyboard_f);
                    return;

            }

            if (message.Text == "Розклад дзвінків")
            {
                ReplyKeyboardMarkup keyboard = new(new[]
                  {
                   new KeyboardButton[] { "Повний" },
                   new KeyboardButton[] { "Скорочений" },
                   new KeyboardButton[] { "Назад" }
                   })
                {
                    ResizeKeyboard = true
                };
                await botClient.SendTextMessageAsync(message.Chat.Id, "Оберіть", replyMarkup: keyboard);
                return;
            }

            if (message.Text == "Повний")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Розклад дзвінків (повний)\r\n1. 08:30 - 09:50\r\n2. 10:00 - 11:20\r\n3. 11:50 - 13:10\r\n4. 13:20 - 14:40\r\n5. 15:10 - 16:30\r\n6. 16:40 - 18:00\r\n7. 18:10 - 19:30\r\n8. 19:40 - 21:00");
                return;
            }

            if (message.Text == "Скорочений")
            {
                await botClient.SendTextMessageAsync(message.Chat.Id, "Розклад дзвінків (скорочений)\r\n1. 08:30 - 09:15\r\n2. 09:25 - 10:10\r\n3. 10:20 - 11:05\r\n4. 11:25 - 12:10\r\n5. 12:20 - 13:05\r\n6. 13:15 - 14:00");
                return;
            }

            if (message.Text == "Назад")
            {
                ReplyKeyboardMarkup keyboard_mm = new(new[]
                {
                new KeyboardButton[] { "Розклад пар"},
                new KeyboardButton[] { "Розклад дзвінків" },
                new KeyboardButton[] { "Вибрати іншу групу" }
                });
            }

            switch (message.Text)
            {
                case "11":
                    await ChangeGroup(update, "11");
                    break;
                case "12":
                    await ChangeGroup(update, "12");
                    break;
                case "13":
                    await ChangeGroup(update, "13");
                    break;
                case "21":
                    await ChangeGroup(update, "21");
                    break;
                case "22":
                    await ChangeGroup(update, "22");
                    break;
                case "23":
                    await ChangeGroup(update, "23");
                    break;
                case "31":
                    await ChangeGroup(update, "31");
                    break;
                case "32":
                    await ChangeGroup(update, "32");
                    break;
                case "41":
                    await ChangeGroup(update, "41");
                    break;
                case "42":
                    await ChangeGroup(update, "42");
                    break;
            }
            if (message.Text == "Вибрати іншу групу")
            {
                ReplyKeyboardMarkup keyboard = new(new[]
                  {
                   new KeyboardButton[] { "А", "Б", "Е" },
                   new KeyboardButton[] { "М", "МД", "МК" },
                   new KeyboardButton[] { "П", "Т", "Ф"  }
                   })
                {
                    ResizeKeyboard = true
                };
                await botClient.SendTextMessageAsync(message.Chat.Id, "Виберіть спеціальність", replyMarkup: keyboard);
                return;
            }

            if (message.Text == "Розклад пар")
            {
                ReplyKeyboardMarkup keyboard = new(new[]
                  {
                   new KeyboardButton[] { "ТИЖДЕНЬ" },
                   new KeyboardButton[] { "ПОНЕДІЛОК", "ВІВТОРОК" },
                   new KeyboardButton[] { "СЕРЕДА", "ЧЕТВЕР"},
                   new KeyboardButton[] { "П'ЯТНИЦЯ"},
                   })
                {
                    ResizeKeyboard = true
                };
                await botClient.SendTextMessageAsync(message.Chat.Id, "Оберіть", replyMarkup: keyboard);
                return;
            }

            if (message.Text == "ТИЖДЕНЬ")
            {
                ReadExcel(GetSpeciality(update), GetGroup(botClient, update), "ТИЖДЕНЬ", botClient, update);
            }

            if (message.Text == "ПОНЕДІЛОК")
            {
                ReadExcel(GetSpeciality(update), GetGroup(botClient, update), "ПОНЕДІЛОК", botClient, update);
            }

            if (message.Text == "ВІВТОРОК")
            {
                ReadExcel(GetSpeciality(update), GetGroup(botClient, update), "ВІВТОРОК", botClient, update);
            }

            if (message.Text == "СЕРЕДА")
            {
                ReadExcel(GetSpeciality(update), GetGroup(botClient, update), "СЕРЕДА", botClient, update);
            }

            if (message.Text == "ЧЕТВЕР")
            {
                ReadExcel(GetSpeciality(update), GetGroup(botClient, update), "ЧЕТВЕР", botClient, update);
            }

            if (message.Text == "П'ЯТНИЦЯ")
            {
                ReadExcel(GetSpeciality(update), GetGroup(botClient, update), "П'ЯТНИЦЯ", botClient, update);
            }

            ReplyKeyboardMarkup keyboard_menu = new(new[]
            {
                new KeyboardButton[] { "Розклад пар"},
                new KeyboardButton[] { "Розклад дзвінків" },
                new KeyboardButton[] { "Вибрати іншу групу" }
                })
            {
                ResizeKeyboard = true
            };
            await botClient.SendTextMessageAsync(message.Chat.Id, "Меню", replyMarkup: keyboard_menu);
        }
        static async Task SynkWithDB(ITelegramBotClient bot, Update update, CancellationToken arg3)
        {
            if (update.Type == UpdateType.Message)
            {
                if (update.Message.Type == MessageType.Text)
                {
                    var chatid = update.Message.Chat.Id;
                    string? username = update.Message.Chat.Username;
                    DateTime.UtcNow.AddHours(2);
                    var date = DateTime.Now.ToString();
                    Console.WriteLine($" {chatid} | {username} | {date}");
                    try
                    {
                        DB = new SQLiteConnection("Data Source = DB.db");
                        DB.Open();
                        SQLiteCommand regcmd = DB.CreateCommand();
                        regcmd.CommandText = "INSERT INTO user(`chatid`, `username`, `date`) VALUES(@chatid, @username, @date)";
                        regcmd.Parameters.AddWithValue("@chatid", chatid.ToString());
                        regcmd.Parameters.AddWithValue("@username", username);
                        regcmd.Parameters.AddWithValue("@date", date);
                        regcmd.ExecuteNonQuery();
                        DB.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex);
                    }
                    await bot.SendTextMessageAsync(update.Message.Chat.Id, "Додано в бд");
                }
            }
        }
        static async Task Audit(ITelegramBotClient bot, Update update, CancellationToken arg3)
        {
            int res = 0;
            try
            {
                string a = "null";
                DB = new SQLiteConnection("Data Source = DB.db");
                DB.Open();
                SQLiteCommand regcmd = DB.CreateCommand();
                regcmd.CommandText = "SELECT IIF((SELECT `chatid` from `user` where username = @username) = @chatid, true, false)";
                regcmd.Parameters.AddWithValue("@chatid", update.Message.Chat.Id.ToString());
                regcmd.Parameters.AddWithValue("@username", update.Message.Chat.Username);
                object result = regcmd.ExecuteScalar();
                res = int.Parse(string.Format("{0}", result));
                regcmd.ExecuteNonQuery();
                DB.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
            if (res == 0)
            {
                await SynkWithDB(bot, update, arg3);
            }
            await bot.SendTextMessageAsync(update.Message.Chat.Id, "Перевірка");
        }
        static async Task ChangeGroup(Update update, string group)
        {
            try
            {
                DB = new SQLiteConnection("Data Source = DB.db");
                DB.Open();
                SQLiteCommand regcmd = DB.CreateCommand();
                regcmd.CommandText = "Update `user` set `group`=@group where chatid=@chatid";
                regcmd.Parameters.AddWithValue("@chatid", update.Message.Chat.Id);
                regcmd.Parameters.AddWithValue("@group", group);
                regcmd.ExecuteNonQuery();
                DB.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }
        static async Task ChangeSpeciality(ITelegramBotClient bot, Update update, string speciality)
        {
            try
            {
                DB = new SQLiteConnection("Data Source = DB.db");
                DB.Open();
                SQLiteCommand regcmd = DB.CreateCommand();
                regcmd.CommandText = "Update `user` set `speciality`=@speciality where chatid=@chatid";
                regcmd.Parameters.AddWithValue("@chatid", update.Message.Chat.Id);
                regcmd.Parameters.AddWithValue("@speciality", speciality);
                regcmd.ExecuteNonQuery();
                DB.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }
        static Task HandleErrorAsync(ITelegramBotClient client, Exception exception, CancellationToken cancellationToken)
        {
            var ErrorMessage = exception switch
            {
                ApiRequestException apiRequestException
                    => $"EROR Telegram API:\n{apiRequestException.ErrorCode}\n{apiRequestException.Message}",
                _ => exception.ToString()
            };
            Console.WriteLine(ErrorMessage);
            return Task.CompletedTask;
        }
        public static int GetCountUsedColumns(Worksheet ws)
        {
            for (int n = 1; ; n++)
            {
                if (String.IsNullOrEmpty((ws.Cells[2, n]).Value))
                    if (String.IsNullOrEmpty((ws.Cells[2, n + 1]).Value))
                        if (String.IsNullOrEmpty((ws.Cells[2, n + 2]).Value))
                            if (String.IsNullOrEmpty((ws.Cells[2, n + 3]).Value))
                            {
                                return n;
                            }
            }
        }
        public static async void ShowDay(string day, int q, Worksheet ws, string rozklad, int i, int n, ITelegramBotClient bot, Update update)
        {
            do
            {
                day = "";
                if (String.IsNullOrEmpty((ws.Cells[q, 1]).Value))
                {
                }
                else
                {
                    day += "\n" + (ws.Cells[q, 1]).Value;
                    //Console.WriteLine(day);
                    //await bot.SendTextMessageAsync(update.Message.Chat.Id, day);
                    rozklad += day + "\n\n";
                }
                if (String.IsNullOrEmpty((ws.Cells[q, i]).Value))
                {
                }
                else
                {
                    if (String.IsNullOrEmpty((ws.Cells[q + 2, i]).Value))
                    {
                        rozklad += (ws.Cells[q, n]).Value + ") " + (ws.Cells[q, i]).Value + "\n\n";
                    }
                    else
                    {
                        if ((ws.Cells[q - 2, n]).Value != (ws.Cells[q + 2, n]).Value)
                        {
                            rozklad += (ws.Cells[q - 2, n]).Value + "(0,5)) " + (ws.Cells[q, i]).Value + "\n\n";
                        }
                        else
                        {
                            rozklad += (ws.Cells[q, n]).Value + "(0,5)) " + (ws.Cells[q, i]).Value + "\n\n";
                            rozklad += (ws.Cells[q, n]).Value + "(0,5)) " + (ws.Cells[q + 2, i]).Value + "\n\n";
                            q += 3;
                        }
                    }
                }
                if (q == 82)
                    break;
                q++;
            }
            while (String.IsNullOrEmpty((ws.Cells[q, 1]).Value));
            Console.WriteLine(rozklad);
            await bot.SendTextMessageAsync(update.Message.Chat.Id, rozklad);
        }
        public static string GetSpeciality(Update update)
        {
            string speciality = "";
            try
            {
                DB = new SQLiteConnection("Data Source = DB.db");
                DB.Open();
                SQLiteCommand regcmd = DB.CreateCommand();
                regcmd.CommandText = "SELECT `speciality` from `user` where chatid=@chatid";
                regcmd.Parameters.AddWithValue("@chatid", update.Message.Chat.Id);
                regcmd.ExecuteNonQuery();
                speciality = (string)regcmd.ExecuteScalar();
                DB.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
            return speciality;
        }
        public static string GetGroup(ITelegramBotClient bot, Update update)
        {
            string group = "";
            try
            {
                DB = new SQLiteConnection("Data Source = DB.db");
                DB.Open();
                SQLiteCommand regcmd = DB.CreateCommand();
                group = regcmd.CommandText = "SELECT `group` from `user` where chatid=@chatid";
                regcmd.Parameters.AddWithValue("@chatid", update.Message.Chat.Id);
                regcmd.ExecuteNonQuery();
                group = (string)regcmd.ExecuteScalar();
                DB.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
            return group;
        }
        public static async void ReadExcel(string speciality, string group, string show, ITelegramBotClient bot, Update update)
        {
            string filePath = @"F:\Visual Studio Projects\Teleg_BOT\rozklad.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook WB = excel.Workbooks.Open(filePath);
            Worksheet ws = (Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range cell = ws.Cells[3, 2];
            string CellValue = cell.Value.ToString();
            string rozklad = "";
            string day = "";
            int countColumn = GetCountUsedColumns(ws);
            int i;//стовпець вибраної групи
            int n = 0;//для запам'ятовування стовпця для номера пари
            int numbRowDayOfWeek = 0;
            Console.OutputEncoding = UTF8Encoding.UTF8;

            for (i = 1; i <= countColumn; i++)
            {

                if (String.IsNullOrEmpty((ws.Cells[2, i]).Value))
                {
                }
                else
                {
                    if (((ws.Cells[2, i]).Value == speciality + " - " + group) || ((ws.Cells[2, i]).Value == speciality + "- " + group) ||
                        ((ws.Cells[2, i]).Value == speciality + "-" + group))
                    {
                        Console.WriteLine(speciality + " - " + group);
                        break;
                    }
                }
            }
            for (int a = i; ; a--)
            {
                if (String.IsNullOrEmpty((ws.Cells[2, a]).Value))
                {
                    if (String.IsNullOrEmpty((ws.Cells[2, a - 1]).Value))
                    {
                        n = a;
                        break;
                    }
                }
            }
            if (show == "ТИЖДЕНЬ")
            {
                for (int j = 3; j <= 82; j++)
                {
                    day = "";
                    if (String.IsNullOrEmpty((ws.Cells[j, 1]).Value))
                    {
                    }
                    else
                    {
                        day += "\n\n" + (ws.Cells[j, 1]).Value + "\n";
                        rozklad += day + "\n";
                    }
                    if (String.IsNullOrEmpty((ws.Cells[j, i]).Value))
                    {
                    }
                    else
                    {
                        if (String.IsNullOrEmpty((ws.Cells[j + 2, i]).Value))
                        {
                            
                            rozklad += (ws.Cells[j, n]).Value + ") " + (ws.Cells[j, i]).Value + "\n\n";
                        }
                        else
                        {
                            if (((ws.Cells[j - 2, n]).Value != (ws.Cells[j + 2, n]).Value))
                            {
                                rozklad += (ws.Cells[j - 2, n]).Value + ")(0,4) " + (ws.Cells[j, i]).Value + "\n";
                            }
                            else
                            {
                                rozklad += (ws.Cells[j, n]).Value + ")(0,5) " + (ws.Cells[j, i]).Value + "\n\n";
                                rozklad += (ws.Cells[j, n]).Value + ")(0,5) " + (ws.Cells[j + 2, i]).Value + "\n\n";
                                j += 3;
                            }
                        }
                    }
                    if (j == 82)
                    {
                        await bot.SendTextMessageAsync(update.Message.Chat.Id, rozklad);
                    }
                }
            }
            else
            {
                for (numbRowDayOfWeek = 1; ; numbRowDayOfWeek++)
                {
                    if ((ws.Cells[numbRowDayOfWeek, 1]).Value == show)
                    {
                        break;
                    }
                }
                ShowDay(day, numbRowDayOfWeek, ws, rozklad, i, n, bot, update);
            }
        }
    }
}
