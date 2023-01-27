//https://github.com/chunfeilung/slt

using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Windows.Interop;
using System.Runtime.InteropServices;



namespace Darpana
{
    public partial class MainWindow : Window
    {
        readonly string OWAPIKEY = Environment.GetEnvironmentVariable("OWAPIKEY", EnvironmentVariableTarget.User); //API key for openweather
        private readonly SpeechRecognitionEngine speechRecognizer = new SpeechRecognitionEngine(); //Global speech recognition engine

        [DllImport("User32")]
        private static extern int ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("User32")]

        private static extern bool SetForegroundWindow(IntPtr hWnd);


        //For Deserializing current weather Json file
        public class WeatherInfo
        {
            [JsonProperty("weather")]
            public DataTable Weather { get; set; }

            [JsonProperty("main")]
            public Dictionary<string, string> Main { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }
        }

        //For Deserializing json objects into a list of dictionaries. Used for Forecast json file.
        public class ForecastInfo
        {
            [JsonProperty("list")]
            public List<Dictionary<string, Object>> ForecastList { get; set; }
        }

        //For Deserializing objects to dictionaries. Used for main portions of forecast such as temps
        public static Dictionary<string, TValue> ToDictionary<TValue>(object obj)
        {
            var json = JsonConvert.SerializeObject(obj);
            var dictionary = JsonConvert.DeserializeObject<Dictionary<string, TValue>>(json);
            return dictionary;
        }

        //For Deserializing objects into a data table. Used for Conditions portion of forecast
        public static DataTable ToDataTable(object obj)
        {
            var json = JsonConvert.SerializeObject(obj);
            var dataTable = JsonConvert.DeserializeObject<DataTable>(json);
            return dataTable;
        }

        //Gets current weather from OpenWeatherMap API Json file
        public async void GetCurrWeather()
        {
            var currWeather = await GetCurrWeatherInfo();//Gets json file from OpenWeatherMap
            var currInfo = JsonConvert.DeserializeObject<WeatherInfo>(currWeather); //Converts Json file into appropriate types
            var iconUrl = new Uri($"http://openweathermap.org/img/wn/{currInfo.Weather.Rows[0][3]}@2x.png"); //Gets Icon image from OpenWeatherMap
            UpdateCurrWeather(currInfo, iconUrl); //Updates the display with new data
        }

        //Updates current weather text blocks in XAML file
        public void UpdateCurrWeather(WeatherInfo currInfo, Uri iconUrl)
        {
            //These filter out the needed text from the new data for the TextBlocks in the XAML file
            var city = currInfo.Name;
            var conditions = currInfo.Weather.Rows[0][1].ToString();
            var temp = currInfo.Main["temp"].Split('.');
            var maxTemp = currInfo.Main["temp_max"].Split('.');
            var minTemp = currInfo.Main["temp_min"].Split('.');

            //This updates the textblocks within the XAML file to show correct data
            City.Text = $"{city}, IA";
            CurrTemp.Text = $"{temp[0]}°F";
            MaxTemp.Text = $"{maxTemp[0]}°F";
            MinTemp.Text = $"{minTemp[0]}°F";
            CurrConditions.Text = conditions;
            WeatherIcon.Source = new BitmapImage(iconUrl);
        }

        //Gets 5 day forecast
        public async void GetForecast()
        {
            var forecastData = await GetForecastInfo(); //Gets results from Json forecast file
            var forecastTime = await GetForecastTime(); //Gets day of forecast
            var forecastInfo = JsonConvert.DeserializeObject<ForecastInfo>(forecastData); //Deserializes json file into appropriate types
            var forecastTimeInfo = JsonConvert.DeserializeObject<ForecastInfo>(forecastTime);

            UpdateForecast(forecastTimeInfo, forecastInfo, 8); //Updates next day forecast
            UpdateForecast(forecastTimeInfo, forecastInfo, 16); //Updates forecast for 2 days later

        }

        //Updates XAML textboxes for forecast with new data. Forecast updates every 3 hours until 5 days is achieved
        public void UpdateForecast(ForecastInfo forecastTimeInfo, ForecastInfo forecastInfo, int hours)
        {
            var days = hours / 8; //Divides hours by 8. This will get either 1 or 2 for days
            var forecastMain = ToDictionary<string>(forecastInfo.ForecastList[days]["temp"]); //Makes main forecast items (temps) a dictionary.
            var forecastWeather = ToDataTable(forecastInfo.ForecastList[days]["weather"]); //Makes weather into a dataTable
            var forecastTempMax = forecastMain["max"].Split('.'); //Filter portion of string out to get integer temp
            var forecastTempMin = forecastMain["min"].Split('.');
            var iconUrl = new Uri($"http://openweathermap.org/img/wn/{forecastWeather.Rows[0][3]}@2x.png"); //Get weather condition icon from OpenWeatherMap
            var dateValue = CreateDate(forecastTimeInfo, hours); //Runs CreateDate method to update day of week

            //If only one day ahead
            if (days == 1)
            {
                //Update Textboxes in XAML file to new data.
                ForecastDT1.Text = dateValue.ToString("ddd"); //Shows day of week abbreviation
                ForecastTempMax1.Text = $"{forecastTempMax[0]}°F";
                ForecastTempMin1.Text = $"{forecastTempMin[0]}°F";
                ForecastIcon1.Source = new BitmapImage(iconUrl);
            }

            //if 2 days ahead
            else
            {
                //Update Textboxes in XAML file to new data
                ForecastDT2.Text = dateValue.ToString("ddd"); //Shows day of week abbreviation
                ForecastTempMax2.Text = $"{forecastTempMax[0]}°F";
                ForecastTempMin2.Text = $"{forecastTempMin[0]}°F";
                ForecastIcon2.Source = new BitmapImage(iconUrl);
            }

        }

        //Takes date out of Json file, splits and parses it to be in correct format for DateTime constructor
        public DateTime CreateDate(ForecastInfo forecastInfo, int time)
        {
            //Splits date text into necessary pieces to then parse for DateTime
            var forecastDate1 = forecastInfo.ForecastList[time]["dt_txt"].ToString();
            var splitDate1 = forecastDate1.Split('-'); //splits year, month, and day apart
            var splitDay1 = splitDate1[2].Split(' '); //splits day from time

            //parses everything to an into to allow for DateTime constructor to work
            Int32.TryParse(splitDate1[0], out var year);
            Int32.TryParse(splitDate1[1], out var month);
            Int32.TryParse(splitDay1[0], out var day);

            var dateValue = new DateTime(year, month, day); //Gets weekday of inserted date
            return dateValue;
        }

        //Starts speech recognition to listen to first command "Darpana"
        public void SpeechRecognition()
        {
            speechRecognizer.SpeechRecognized += SpeechRecognizer_Speech; //Runs the command heard
            var grammar = new GrammarBuilder();

            var initial = new Choices("darpana"); //Initial command for taking in commands
            var actions = new Choices("hide", "show", "operation"); // all potential actions that can be made
            var modules = new Choices("weather", "time", "everything", "ironhorse"); //all potential modules that can be changed

            //append all the choices to the grammar
            grammar.Append(initial);
            grammar.Append(actions);
            grammar.Append(modules);

            speechRecognizer.LoadGrammar(new Grammar(grammar)); //Loads grammar into recognition engine asynchronously so the engine knows what to listen for
            speechRecognizer.SetInputToDefaultAudioDevice(); //Sets input to default audio (device mic)
            speechRecognizer.RecognizeAsync(RecognizeMode.Multiple); //Recognize speech for multiple words

        }

        //Essenitally same thing as above but for commands
        private void SpeechRecognizer_Speech(object sender, SpeechRecognizedEventArgs e)
        {
            Debug.WriteLine(e.Result.Words.Count);
            if (e.Result.Words.Count == 3) //If there are three words that are said
            {
                //Make words lower case
                var initial = e.Result.Words[0].Text.ToLower();
                var action = e.Result.Words[1].Text.ToLower();
                var module = e.Result.Words[2].Text.ToLower();

                switch (initial)
                {
                    case "darpana":

                        switch (action)
                        {
                            case "hide": //if user said hide, check for next command
                                switch (module)
                                {
                                    //hide module based on if user said weather, time, or everything
                                    case "weather":
                                        Debug.WriteLine("Weather");
                                        WeatherModule.Visibility = Visibility.Hidden;
                                        break;

                                    case "time":
                                        TimeModule.Visibility = Visibility.Hidden;
                                        break;

                                    case "everything":
                                        DarpanaWindow.Visibility = Visibility.Hidden;
                                        break;
                                }
                                break;

                            case "show":
                                switch (module) //if user said show, check for next command
                                {
                                    //show weather, time, or everything if that is what the user said
                                    case "weather":
                                        WeatherModule.Visibility = Visibility.Visible;
                                        break;

                                    case "time":
                                        TimeModule.Visibility = Visibility.Visible;
                                        break;

                                    case "everything":
                                        DarpanaWindow.Visibility = Visibility.Visible;
                                        TimeModule.Visibility = Visibility.Visible;
                                        WeatherModule.Visibility = Visibility.Visible;
                                        break;
                                }
                                break;

                            case "operation":
                                switch (module)
                                {

                                    case "ironhorse": //if user says "darpana operation ironhorse, runs little easteregg

                                        //creates and opens new cmd line then runs command. Closes cmd line once cmd is over
                                        Process process = new Process();
                                        ProcessStartInfo startInfo = new ProcessStartInfo
                                        {
                                            FileName = "cmd.exe",
                                            Arguments = "/C start C:\\users\\braid\\Source\\repos\\braidynbritt\\Darpana\\Darpana\\slt --filled --colored --speed 4"
                                        };

                                        process.StartInfo = startInfo;
                                        process.Start();
                                        process.WaitForExit();

                                        //Resets Darpana window to make sure it is not downsized after ironhorse is ran
                                        WindowState = WindowState.Minimized;
                                        var mainWindow = Application.Current.MainWindow;
                                        ShowWindow(new WindowInteropHelper(mainWindow).Handle, 9);
                                        SetForegroundWindow(new WindowInteropHelper(mainWindow).Handle);
                                        break;
                                }
                                break;
                        }
                        break;
                }

                //Update display
                CommandManager.InvalidateRequerySuggested();
            }
        }

        //Starts the dispatcher timers for updating XAML file/display
        public void StartTimers()
        {
            //DispatcherTimer for Time event. Updates every 500 milliseconds
            var dispatcherTimerTime = new DispatcherTimer();
            dispatcherTimerTime.Tick += new EventHandler(DispatcherTimer_Time);
            dispatcherTimerTime.Interval = new TimeSpan(0, 0, 0, 0, 500);
            dispatcherTimerTime.Start();

            //DispatcherTimer for current weather event. Updates every 30 minutes.
            var dispatcherTimerWeather = new DispatcherTimer();
            dispatcherTimerWeather.Tick += new EventHandler(DispatcherTimer_Weather);
            dispatcherTimerWeather.Interval = new TimeSpan(0, 0, 1800);
            dispatcherTimerWeather.Start();

            //DispatcherTimer for forecast event. Updates every 3 hours.
            var dispatcherTimerForecast = new DispatcherTimer();
            dispatcherTimerWeather.Tick += new EventHandler(DispatcherTimer_Forecast);
            dispatcherTimerWeather.Interval = new TimeSpan(0, 0, 10800);
            dispatcherTimerWeather.Start();
        }

        //Asyn class that does OpenWeatherMap API call. If succesful response is returned then give content to other primary methods
        public async Task<string> GetCurrWeatherInfo()
        {
            HttpClient client = new HttpClient
            {
                BaseAddress = new Uri("http://pro.openweathermap.org/data/2.5/") //API URL for client
            };

            var query = $"weather?lat=42.6411&lon=-95.2097&APPID={OWAPIKEY}&units=imperial"; //Api specifics. Gets location of Storm Lake, IA and gives Farhenheit temp.
            var response = await client.GetAsync(query);
            response.EnsureSuccessStatusCode();
            //If response is successful move on
            return await response.Content.ReadAsStringAsync();

        }

        //Same thing as above method. Runs different API query.
        public async Task<string> GetForecastInfo()
        {
            HttpClient client = new HttpClient
            {
                BaseAddress = new Uri("http://pro.openweathermap.org/data/2.5/")
            };

            var query = $"forecast/daily?lat=42.6411&lon=-95.2097&APPID={OWAPIKEY}&units=imperial";
            var response = await client.GetAsync(query);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsStringAsync();

        }

        //Same as about method. Runs different API
        public async Task<string> GetForecastTime()
        {
            HttpClient client = new HttpClient
            {
                BaseAddress = new Uri("http://pro.openweathermap.org/data/2.5/")
            };

            var query = $"forecast?lat=42.6411&lon=-95.2097&APPID={OWAPIKEY}&units=imperial";
            var response = await client.GetAsync(query);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsStringAsync();
        }

        //DispatcherTimer method for time module. Updated Time and Date
        private void DispatcherTimer_Time(object sender, EventArgs e)
        {
            //Updates time and date text boxes
            Time.Text = DateTime.Now.ToString("T");
            Date.Text = DateTime.Now.ToString("D");

            // Forcing the CommandManager to raise the RequerySuggested event. Updates display
            CommandManager.InvalidateRequerySuggested();
        }

        //DispatcherTimer method for weather
        private void DispatcherTimer_Weather(object sender, EventArgs e)
        {
            // Runs all current weather methods.
            GetCurrWeather();

            // Forcing the CommandManager to raise the RequerySuggested event. Updates display
            CommandManager.InvalidateRequerySuggested();
        }

        //DispatchTimer method for Forecasts
        private void DispatcherTimer_Forecast(object sender, EventArgs e)
        {
            // Runs all forecast methods
            GetForecast();

            // Forcing the CommandManager to raise the RequerySuggested event. Updates display
            CommandManager.InvalidateRequerySuggested();
        }

        //Main function. Runs initial methods
        public MainWindow()
        {
            InitializeComponent(); //WPF method
            GetCurrWeather(); //Runs all weather methods to get current weather on startup
            GetForecast(); //Runs all forecast methods to get forecast on startup
            StartTimers(); //Starts all the dispatcher timers for updating display
            SpeechRecognition(); //Start speech recognition

            Time.Text = DateTime.Now.ToString("T"); //Gets current time on startup and sets appropriate text block
            Date.Text = DateTime.Now.ToString("D"); //Gets current date on startup and sets appropriate text block
        }
    }
}