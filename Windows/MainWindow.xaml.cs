using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using WpfApp2.HelperClasses;
using WpfApp2.ModelClasses;

namespace WpfApp2
{
    public partial class MainWindow : Window
    {
        private ModelEF model;
        private List<Users> users;
        private List<Auto> autos;
        public MainWindow()
        {
            InitializeComponent();
            model = new ModelEF();
            users = new List<Users>();
            autos = new List<Auto>();
        }
        private void ComboLoadData()

        {
            comboBoxUsers.Items.Clear();
            users = model.Users.ToList();
            foreach (var item in users)
                comboBoxUsers.Items.Add($"{item.FullName} {item.PSeria} {item.PNumber}");
            comboBoxUsers.SelectedIndex = 0;
            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            comboBoxAutos.Items.Clear();
            foreach (var item in autos)
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN} ");
            comboBoxAutos.SelectedIndex = 0;
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ComboLoadData();
        }
        private void comboBoxUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)

        {
            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            comboBoxAutos.Items.Clear();
            foreach (var item in autos)
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN} ");
            comboBoxAutos.SelectedIndex = 0;
        }
        private void SaveDocument_Click(object sender, RoutedEventArgs e)
        {

            // Создаем диалоговое окно для выбора директории
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            // Задаем описание для окна выбора директории
            fbd.Description = "Выберите место сохранения";

            // Проверяем результат открытия диалогового окна
            if (System.Windows.Forms.DialogResult.OK == fbd.ShowDialog())
            {
                // Получаем активного пользователя
                Users activeUser = users[comboBoxUsers.SelectedIndex];
                // Получаем активный автомобиль
                Auto activeAuto = activeUser.Auto.ToList()[comboBoxAutos.SelectedIndex];
                // Создаем документ и сохраняем его в указанную директорию
                CreateDocument(
                    $@"{fbd.SelectedPath}\Купля-Продажа-Автомобиля-{activeUser.FullName}.docx",
                    activeUser,
                    activeAuto);

                // Выводим сообщение о сохранении файла
                System.Windows.MessageBox.Show("Файл сохранён");
            }
        }
        // Метод создания документа с подстановкой данных пользователя и автомобиля
        private void CreateDocument(string directorypath, Users users, Auto auto)
        {
            // Получаем текущую дату
            var today = DateTime.Now.ToShortDateString();

            // Создаем объект для работы с документом Word
            WordHelper word = new WordHelper("Contractsale.docx");

            // Создаем словарь для замены ключевых слов в документе
            var items = new Dictionary<string, string>
            {
                // Замена ключевого слова <Today> на текущую дату
                {"<Today>", today },
                // Данные пользователя
                { "<FullName>", users.FullName }, // Фио
                {   "<DateOfBirth>",  users.DateOfBirth.Value.ToShortDateString() }, // Дата рождения
                {"<Adress>", users.Adress }, // Адрес
                {"<PSeria>", users.PSeria.ToString() }, // Серия паспорта
                {"<PNumber>", users.PNumber.ToString() }, // Номер паспорта
                {"<PVidan>", users.PVidan }, // Кем выдан паспорт
                // Данные автомобиля
                { "<ModelV>", auto.Model }, // Модель автомобиля
                { "<CategoryV>", auto.Category }, // Категория автомобиля
                { "<TypeV>", auto.TypeV }, // Тип автомобиля
                { "<VIN>", auto.VIN }, // VIN номер
                { "<RegistrationMark>", auto.RegistrationMark }, // Регистрационный знак 
                {"<YearV>", auto.YearOfRelease.Value.Year. ToString() }, // Год выпуска 
                {"<EngineV>", auto.EngineNumber }, // Номер двигателя
                {"<ChassisV>", auto.Chassis }, // Шасси
                {"<BodyworkV>", auto.Bodywork },  // Кузов
                {"<ColorV>", auto.Color }, // Цвет
                {"<SeriaPV>", auto.SeriaPasport }, // Серия ПТС
                {"<NumberPV>", auto.NumbePasport },  // Номер ПТС
                {"<VidanPV>", auto.VidanPasport } // Кем выдан ПТС
            };
                // Обрабатывает документ, подставляя значения из словаря вместо ключевых слов
                word.Process(items, directorypath);
        }
    }
}
