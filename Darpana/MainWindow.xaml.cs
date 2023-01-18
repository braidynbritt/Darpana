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

namespace Darpana
{
    public partial class MainWindow : Window
    {
        //Global API key
        string OWAPIKEY;

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
        public void GetCurrWeather()
        {
            var currWeather = GetCurrWeatherInfo().Result; //Gets json file from OpenWeatherMap
            var currInfo = JsonConvert.DeserializeObject<WeatherInfo>(currWeather); //Converts Json file into appropriate types
            var iconUrl = new Uri(String.Concat("http://openweathermap.org/img/wn/", currInfo.Weather.Rows[0][3].ToString(), "@2x.png")); //Gets Icon image from OpenWeatherMap
            UpdateCurrWeather(currInfo, iconUrl); //Updates the display with new data

        }
        
        //Updates current weather text blocks in XAML file
        public void UpdateCurrWeather(WeatherInfo currInfo, Uri iconUrl)
        {
            //These filter out the needed text from the new data for the TextBlocks in the XAML file
            string city = currInfo.Name.ToString();
            string conditions = currInfo.Weather.Rows[0][1].ToString();
            string[] temp = currInfo.Main["temp"].ToString().Split('.');
            string[] maxTemp = currInfo.Main["temp_max"].ToString().Split('.');
            string[] minTemp = currInfo.Main["temp_min"].ToString().Split('.');

            //This updates the textblocks within the XAML file to show correct data
            City.Text     = $"{city}, IA";
            CurrTemp.Text = $"{temp[0]}, °F";
            MaxTemp.Text  = $"{maxTemp[0]}, °F";
            MinTemp.Text  = $"{minTemp[0]}, °F";
            CurrConditions.Text = conditions;
            WeatherIcon.Source = new BitmapImage(iconUrl);
        }

        //Gets 5 day forecast
        public void GetForecast()
        {
            string forecastData = GetForecastInfo().Result; //Gets results from Json forecast file
            ForecastInfo forecastInfo = JsonConvert.DeserializeObject<ForecastInfo>(forecastData); //Deserializes json file into appropriate types

            UpdateForecast(forecastInfo, 1); //Updates next day forecast
            UpdateForecast(forecastInfo, 2); //Updates forecast for 2 days later

        }

        //Updates XAML textboxes for forecast with new data. Forecast updates every 3 hours until 5 days is achieved
        public void UpdateForecast(ForecastInfo forecastInfo, int days)
        {
            int hours = days * 8; //Multiply days by 8. This will get either 8 or 16. The 3 hours multipled by 8 or 16 gives 24 or 48 hours ahead
            Dictionary<string, string> forecastMain = ToDictionary<string>(forecastInfo.ForecastList[days]["main"]); //Makes main forecast items (temps) a dictionary.
            DataTable forecastWeather = ToDataTable(forecastInfo.ForecastList[days]["weather"]); //Makes weather into a dataTable
            string[] forecastTempMax = forecastMain["temp_max"].ToString().Split('.'); //Filter portion of string out to get integer temp
            string[] forecastTempMin = forecastMain["temp_min"].ToString().Split('.');
            Uri iconUrl = new Uri(String.Concat("http://openweathermap.org/img/wn/", forecastWeather.Rows[0][3].ToString(), "@2x.png")); //Get weather condition icon from OpenWeatherMap
            DateTime dateValue = CreateDate(forecastInfo, hours); //Runs CreateDate method to update day of week


            //If only one day ahead
            if (days == 1)
            {
                //Update Textboxes in XAML file to new data.
                ForecastDT1.Text = dateValue.ToString("ddd"); //Shows day of week abbreviation
                ForecastTempMax1.Text = String.Concat(forecastTempMax[0], "°F");
                ForecastTempMin1.Text = String.Concat(forecastTempMin[0], "°F");
                ForecastIcon1.Source = new BitmapImage(iconUrl);
            }

            //if 2 days ahead
            else
            {
                //Update Textboxes in XAML file to new data
                ForecastDT2.Text = dateValue.ToString("ddd"); //Shows day of week abbreviation
                ForecastTempMax2.Text = String.Concat(forecastTempMax[0], "°F");
                ForecastTempMin2.Text = String.Concat(forecastTempMin[0], "°F");
                ForecastIcon2.Source = new BitmapImage(iconUrl);
            }

        }

        //Takes date out of Json file, splits and parses it to be in correct format for DateTime constructor
        public DateTime CreateDate(ForecastInfo forecastInfo, int time)
        {
            //Splits date text into necessary pieces to then parse for DateTime
            string forecastDate1 = forecastInfo.ForecastList[time]["dt_txt"].ToString();
            string[] splitDate1 = forecastDate1.Split('-'); //splits year, month, and day apart
            string[] splitDay1 = splitDate1[2].Split(' '); //splits day from time

            //parses everything to an into to allow for DateTime constructor to work
            Int32.TryParse(splitDate1[0], out int year);
            Int32.TryParse(splitDate1[1], out int month);
            Int32.TryParse(splitDay1[0], out int day);

            DateTime dateValue = new DateTime(year, month, day); //Gets weekday of inserted date
            return dateValue;
        }

        //Starts the dispatcher timers for updating XAML file/display
        public void StartTimers()
        {
            //DispatcherTimer for Time event. Updates every 500 milliseconds
            DispatcherTimer dispatcherTimerTime = new DispatcherTimer();
            dispatcherTimerTime.Tick += new EventHandler(DispatcherTimer_Time);
            dispatcherTimerTime.Interval = new TimeSpan(0, 0, 0, 0, 500);
            dispatcherTimerTime.Start();
            
            //DispatcherTimer for current weather event. Updates every 30 minutes.
            DispatcherTimer dispatcherTimerWeather = new DispatcherTimer();
            dispatcherTimerWeather.Tick += new EventHandler(DispatcherTimer_Weather);
            dispatcherTimerWeather.Interval = new TimeSpan(0, 0, 1800);
            dispatcherTimerWeather.Start();

            //DispatcherTimer for forecast event. Updates every 3 hours.
            DispatcherTimer dispatcherTimerForecast = new DispatcherTimer();
            dispatcherTimerWeather.Tick += new EventHandler(DispatcherTimer_Forecast);
            dispatcherTimerWeather.Interval = new TimeSpan(0, 0, 10800);
            dispatcherTimerWeather.Start();
        }

        //Asyn class that does OpenWeatherMap API call. If succesful response is returned then give content to other primary methods
        public async Task<string> GetCurrWeatherInfo()
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://pro.openweathermap.org/data/2.5/"); //API URL for client
            var query = $"weather?lat=42.6411&lon=-95.2097&APPID={OWAPIKEY}&units=imperial"; //Api specifics. Gets location of Storm Lake, IA and gives Farhenheit temp.
            HttpResponseMessage response = client.GetAsync(query).Result;
            response.EnsureSuccessStatusCode(); //If response is successful move on
            string s = await response.Content.ReadAsStringAsync(); //Do other things while waiting for response
            return s;

        }

        //Same thing as above class. Runs different API query.
        public async Task<string> GetForecastInfo()
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("http://pro.openweathermap.org/data/2.5/");
            var query = $"forecast?lat=42.6411&lon=-95.2097&APPID={OWAPIKEY}&units=imperial";
            HttpResponseMessage response = client.GetAsync(query).Result;
            response.EnsureSuccessStatusCode();
            string s = await response.Content.ReadAsStringAsync();
            return s;
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
            //Gets API key from user environment 
            OWAPIKEY = Environment.GetEnvironmentVariable("OWAPIKEY", EnvironmentVariableTarget.User);

            InitializeComponent(); //WPF method
            Time.Text = DateTime.Now.ToString("T"); //Gets current time on startup and sets appropriate text block
            Date.Text = DateTime.Now.ToString("D"); //Gets current date on startup and sets appropriate text block
            GetCurrWeather(); //Runs all weather methods to get current weather on startup
            GetForecast(); //Runs all forecast methods to get forecast on startup
            StartTimers(); //Starts all the dispatcher timers for updating display
        }

    }
}
