using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ical.Net;
using Ical.Net.DataTypes;
using Ical.Net.Serialization.iCalendar.Serializers;

namespace ExtractCalendar
{
    internal class Program
    {
        private static readonly CookieContainer CookieContainer = new CookieContainer();

        private static readonly HttpClientHandler Handler = new HttpClientHandler
        {
            CookieContainer = CookieContainer,
            AutomaticDecompression =
                DecompressionMethods.GZip | DecompressionMethods.Deflate | DecompressionMethods.Brotli
        };

        private static readonly HttpClient Client = new HttpClient(Handler);
        private static string ViewState { get; set; }

        private static async Task Main(string[] args)
        {
            if (args.Length != 3) throw new ArgumentException("Wrong arguments, should be: username password path");

            Client.BaseAddress = new Uri("https://aurionweb.ensiie.fr/");
            Client.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36");
            Client.DefaultRequestHeaders.Add("Accept", "text/html");
            Client.DefaultRequestHeaders.Add("Accept-Encoding", "br");
            Client.DefaultRequestHeaders.Add("Accept-Language", "en-GB");
            Client.DefaultRequestHeaders.Add("Sec-Fetch-Mode", "navigate");
            Client.DefaultRequestHeaders.Add("Sec-Fetch-User", "?1");
            Client.DefaultRequestHeaders.Add("Sec-Fetch-Site", "same-origin");
            Client.DefaultRequestHeaders.Add("Connection", "keep-alive");
            Client.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");
            Client.DefaultRequestHeaders.Add("Cache-Control", "max-age=0");
            Client.DefaultRequestHeaders.Add("DNT", "1");

            foreach (var v in Client.DefaultRequestHeaders.Accept)
                if (v.MediaType.Contains("text/html"))
                {
                    var field = v.GetType().GetTypeInfo().BaseType
                        ?.GetField("_mediaType", BindingFlags.NonPublic | BindingFlags.Instance);
                    field?.SetValue(v,
                        "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3");
                    v.CharSet = "";
                }

            foreach (var v in Client.DefaultRequestHeaders.AcceptLanguage)
                if (v.Value.Contains("en-GB"))
                {
                    var field = v.GetType().GetField("_value", BindingFlags.NonPublic | BindingFlags.Instance);
                    field.SetValue(v, "en-GB,en;q=0.9,en-US;q=0.8,fr;q=0.7");
                }

            try
            {
                await ConnectToAurion(args[0], args[1]);

                foreach (var user in new List<string>(ConfigurationManager.AppSettings["users"].Split(new[] {';'})))
                    if (!string.IsNullOrEmpty(user))
                    {
                        var (responseBody, name) = await GetCalendar(user);

                        var jsonBody = ParseXml(responseBody);

                        var listCalendarEvent = ParseJson(jsonBody);

                        ExportToIcs(listCalendarEvent, args[2] + name + ".ics");
                    }
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine("\nException Caught!");
                Console.WriteLine("Message :{0} ", e.Message);
            }
        }

        private static async Task ConnectToAurion(string username, string password)
        {
            var response = await Client.GetAsync("/");
            response.EnsureSuccessStatusCode();
            var responseBody = await response.Content.ReadAsStringAsync();
            Client.DefaultRequestHeaders.Add("Referer",
                "https://cas.ensiie.fr/login?service=https%3A%2F%2Faurionweb.ensiie.fr%2F%2Flogin%2Fcas");
            Client.DefaultRequestHeaders.Add("Origin", "https://cas.ensiie.fr");
            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("username", username),
                new KeyValuePair<string, string>("password", password),
                new KeyValuePair<string, string>("execution",
                    Regex.Match(responseBody,
                            @"<input type=""hidden"" name=""execution"" value=""(.+?)""/><input type=""hidden""")
                        .Groups[1]
                        .Value),
                new KeyValuePair<string, string>("_eventId", "submit"),
                new KeyValuePair<string, string>("geolocation", "")
            });
            var loginResult = Client
                .PostAsync("https://cas.ensiie.fr/login?service=https%3A%2F%2Faurionweb.ensiie.fr%2F%2Flogin%2Fcas",
                    content)
                .Result;
            loginResult.EnsureSuccessStatusCode();
            responseBody = await loginResult.Content.ReadAsStringAsync();
            ViewState = GetViewStateFromBody(responseBody);
        }

        private static void ExportToIcs(List<CalendarEvent> listCalendarEvent, string path)
        {
            var calendar = new Calendar();
            calendar.AddTimeZone(new VTimeZone("Europe/Paris"));
            Ical.Net.CalendarEvent e;
            foreach (var calendarEvent in listCalendarEvent)
            {
                e = new Ical.Net.CalendarEvent
                {
                    Start = new CalDateTime(Convert.ToDateTime(calendarEvent.Start)),
                    End = new CalDateTime(Convert.ToDateTime(calendarEvent.End)),
                    IsAllDay = calendarEvent.AllDay,
                    Summary = calendarEvent.Title
                };
                calendar.Events.Add(e);
            }

            var serializer = new CalendarSerializer();
            var serializedCalendar = serializer.SerializeToString(calendar);

            using (var file = new StreamWriter(path))
            {
                file.Write(serializedCalendar);
            }
        }

        private static List<CalendarEvent> ParseJson(string jsonBody)
        {
            var listCalendarEvent = (List<CalendarEvent>) JsonSerializer.Deserialize(jsonBody,
                typeof(List<CalendarEvent>), new JsonSerializerOptions
                {
                    IgnoreNullValues = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
            return listCalendarEvent;
        }

        private static string ParseXml(string responseBody)
        {
            var xDoc = XDocument.Parse(responseBody);
            var jsonBody = xDoc.Descendants("changes").FirstOrDefault()?.Elements("update")
                .First(x => x.FirstAttribute.Value == "form:j_idt114")
                .Value;
            jsonBody = Regex.Match(jsonBody ?? throw new InvalidOperationException(), @"^\{""events"" : (.+?)\}$")
                .Groups[1].Value;
            return jsonBody;
        }

        private static async Task<(string, string)> GetCalendar(string calendarId)
        {
            FormUrlEncodedContent content;
            HttpResponseMessage response;
            string responseBody;
            Client.DefaultRequestHeaders.Remove("Referer");
            Client.DefaultRequestHeaders.Remove("Origin");
            Client.DefaultRequestHeaders.Add("Referer", "https://aurionweb.ensiie.fr/");
            Client.DefaultRequestHeaders.Add("Origin", "https://aurionweb.ensiie.fr");
            response = await Client.GetAsync("/");
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            ViewState = GetViewStateFromBody(responseBody);

            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("javax.faces.partial.ajax", "true"),
                new KeyValuePair<string, string>("javax.faces.source", "form:j_idt50"),
                new KeyValuePair<string, string>("javax.faces.partial.execute", "form:j_idt50"),
                new KeyValuePair<string, string>("javax.faces.partial.render", "form:sidebar"),
                new KeyValuePair<string, string>("form:j_idt50", "form:j_idt50"),
                new KeyValuePair<string, string>("webscolaapp.Sidebar.ID_SUBMENU", "submenu_55642"),
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1603"),
                new KeyValuePair<string, string>("form:sauvegarde", ""),
                new KeyValuePair<string, string>("javax.faces.ViewState", ViewState)
            });
            response = Client.PostAsync("/faces/MainMenuPage.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();

            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1603"),
                new KeyValuePair<string, string>("form:sauvegarde", ""),
                new KeyValuePair<string, string>("javax.faces.ViewState", ViewState),
                new KeyValuePair<string, string>("form:sidebar", "form:sidebar"),
                new KeyValuePair<string, string>("form:sidebar_menuid", "3_1")
            });
            response = Client.PostAsync("/faces/MainMenuPage.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            ViewState = GetViewStateFromBody(responseBody);
            Client.DefaultRequestHeaders.Remove("Referer");
            Client.DefaultRequestHeaders.Add("Referer", "https://aurionweb.ensiie.fr/faces/ChoixPlanning.xhtml");
            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1603"),
                new KeyValuePair<string, string>("form:search-texte", ""),
                new KeyValuePair<string, string>("form:search-texte-avancer", ""),
                new KeyValuePair<string, string>("form:input-expression-exacte", ""),
                new KeyValuePair<string, string>("form:input-un-des-mots", ""),
                new KeyValuePair<string, string>("form:input-aucun-des-mots", ""),
                new KeyValuePair<string, string>("form:input-nombre-debut", ""),
                new KeyValuePair<string, string>("form:input-nombre-fin", ""),
                new KeyValuePair<string, string>("form:calendarDebut_input", ""),
                new KeyValuePair<string, string>("form:calendarFin_input", ""),
                new KeyValuePair<string, string>("form:j_idt185_reflowDD", "0_0"),
                new KeyValuePair<string, string>("form:j_idt185:j_idt190:filter", ""),
                new KeyValuePair<string, string>("form:j_idt185:j_idt192:filter", ""),
                new KeyValuePair<string, string>("form:j_idt185_checkbox", "on"),
                new KeyValuePair<string, string>("form:j_idt185_selection", calendarId),
                new KeyValuePair<string, string>("form:j_idt248", ""),
                new KeyValuePair<string, string>("javax.faces.ViewState", ViewState)
            });
            response = Client.PostAsync("/faces/ChoixPlanning.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();

            var name = GetNameFromBody(responseBody);
            ViewState = GetViewStateFromBody(responseBody);
            Client.DefaultRequestHeaders.Remove("Referer");
            Client.DefaultRequestHeaders.Add("Referer", "https://aurionweb.ensiie.fr/faces/Planning.xhtml");

            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("javax.faces.partial.ajax", "true"),
                new KeyValuePair<string, string>("javax.faces.source", "form:j_idt114"),
                new KeyValuePair<string, string>("javax.faces.partial.execute", "form:j_idt114"),
                new KeyValuePair<string, string>("form:j_idt114", "form:j_idt114"),
                new KeyValuePair<string, string>("javax.faces.partial.render", "form:j_idt114"),
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>("form:j_idt114_view", "month"),
                new KeyValuePair<string, string>("form:offsetFuseauNavigateur", "-7200000"),
                new KeyValuePair<string, string>("javax.faces.ViewState", ViewState),
                new KeyValuePair<string, string>("form:calendarDebut_input", ""),
                new KeyValuePair<string, string>("form:j_idt114_end", "1570831200000"),
                new KeyValuePair<string, string>("form:j_idt114_start", "1567375200000"),
                new KeyValuePair<string, string>("form:onglets_activeIndex", "0"),
                new KeyValuePair<string, string>("form:onglets_scrollState", "0"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1620")
            });
            response = Client.PostAsync("/faces/Planning.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            return (responseBody, name);
        }

        private static string GetViewStateFromBody(string body)
        {
            return Regex.Match(body, @"id=""j_id1:javax.faces.ViewState:0"" value=""(.+?)"" autocomplete=""off""")
                .Groups[1].Value;
        }

        private static string GetNameFromBody(string body)
        {
            return Regex.Match(body,
                    @"class=""ui-messages ui-widget"" aria-live=""polite""></div><div class=""divCurrentUser"">(.+?)</div>")
                .Groups[1].Value;
        }
    }
}