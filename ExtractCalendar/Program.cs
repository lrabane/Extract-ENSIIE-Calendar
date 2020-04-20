using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;

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

        private static async Task Main(string[] args)
        {
            if (args.Length != 3) throw new ArgumentException("Wrong arguments, should be: username password path");

            Client.BaseAddress = new Uri("https://aurionweb.ensiie.fr/");
            Client.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; rv:68.0) Gecko/20100101 Firefox/68.0");
            Client.DefaultRequestHeaders.Add("Accept", "text/html");
            Client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate, br");
            Client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.5");
            Client.DefaultRequestHeaders.Add("Sec-Fetch-Mode", "navigate");
            Client.DefaultRequestHeaders.Add("Sec-Fetch-User", "?1");
            Client.DefaultRequestHeaders.Add("Sec-Fetch-Site", "same-origin");
            Client.DefaultRequestHeaders.Add("Connection", "keep-alive");
            Client.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");
            Client.DefaultRequestHeaders.Add("Cache-Control", "max-age=0");
            Client.DefaultRequestHeaders.Add("DNT", "1");

            try
            {
                await ConnectToAurion(args[0], args[1]);

                foreach (var user in new List<string>(ConfigurationManager.AppSettings["users"].Split(new[] {';'})).Where(user => !string.IsNullOrEmpty(user)))
                {
                    var (body, name, calendarForm) = await GetCalendar(user);

                    var jsonBody = ParseJsonBody(body, calendarForm);

                    var listCalendarEvent = DeserializeCalendar(jsonBody);

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
            var execution = ParseExecution(responseBody);

            Client.DefaultRequestHeaders.Add("Referer",
                "https://cas.ensiie.fr/login?service=https%3A%2F%2Faurionweb.ensiie.fr%2F%2Flogin%2Fcas");
            Client.DefaultRequestHeaders.Add("Origin", "https://cas.ensiie.fr");
            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("username", username),
                new KeyValuePair<string, string>("password", password),
                new KeyValuePair<string, string>("execution", execution),
                new KeyValuePair<string, string>("_eventId", "submit"),
                new KeyValuePair<string, string>("geolocation", ""),
                new KeyValuePair<string, string>("submit", "LOGIN")
            });
            var loginResult = Client
                .PostAsync("https://cas.ensiie.fr/login?service=https%3A%2F%2Faurionweb.ensiie.fr%2F%2Flogin%2Fcas",
                    content)
                .Result;
            loginResult.EnsureSuccessStatusCode();
        }


        private static async Task<(string, string, string)> GetCalendar(string calendarId)
        {
            Client.DefaultRequestHeaders.Remove("Referer");
            Client.DefaultRequestHeaders.Remove("Origin");
            Client.DefaultRequestHeaders.Add("Referer", "https://aurionweb.ensiie.fr/");
            Client.DefaultRequestHeaders.Add("Origin", "https://aurionweb.ensiie.fr");
            var response = await Client.GetAsync("/");
            response.EnsureSuccessStatusCode();
            var responseBody = await response.Content.ReadAsStringAsync();
            var viewState = ParseViewState(responseBody);
            var mainMenuForm = ParseMainMenuForm(responseBody);
            var calendarMenuForm = ParseCalendarMenuForm(responseBody);

            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("javax.faces.partial.ajax", "true"),
                new KeyValuePair<string, string>("javax.faces.source", mainMenuForm),
                new KeyValuePair<string, string>("javax.faces.partial.execute", mainMenuForm),
                new KeyValuePair<string, string>("javax.faces.partial.render", "form:sidebar"),
                new KeyValuePair<string, string>(mainMenuForm, mainMenuForm),
                new KeyValuePair<string, string>("webscolaapp.Sidebar.ID_SUBMENU", calendarMenuForm),
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1603"),
                new KeyValuePair<string, string>("form:sauvegarde", ""),
                new KeyValuePair<string, string>("javax.faces.ViewState", viewState)
            });
            response = Client.PostAsync("/faces/MainMenuPage.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            var calendarSubMenuId = ParseCalendarSubMenuId(responseBody);

            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1603"),
                new KeyValuePair<string, string>("form:sauvegarde", ""),
                new KeyValuePair<string, string>("javax.faces.ViewState", viewState),
                new KeyValuePair<string, string>("form:sidebar", "form:sidebar"),
                new KeyValuePair<string, string>("form:sidebar_menuid", calendarSubMenuId)
            });
            response = Client.PostAsync("/faces/MainMenuPage.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            viewState = ParseViewState(responseBody);
            var selectCalendarForm = ParseSelectCalendarForm(responseBody);
            var selectCalendarButtonForm = ParseSelectCalendarButtonForm(responseBody);

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
                new KeyValuePair<string, string>(selectCalendarForm + "_reflowDD", "0_0"),
                //new KeyValuePair<string, string>(selectCalendarForm + ":j_idt185:filter", ""),
                //new KeyValuePair<string, string>(selectCalendarForm + ":j_idt187:filter", ""),
                new KeyValuePair<string, string>(selectCalendarForm + "_checkbox", "on"),
                new KeyValuePair<string, string>(selectCalendarForm + "_selection", calendarId),
                new KeyValuePair<string, string>(selectCalendarButtonForm, ""),
                new KeyValuePair<string, string>("javax.faces.ViewState", viewState)
            });
            response = Client.PostAsync("/faces/ChoixPlanning.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            viewState = ParseViewState(responseBody);
            var calendarForm = ParseCalendarForm(responseBody);
            var name = ParseName(responseBody);

            Client.DefaultRequestHeaders.Remove("Referer");
            Client.DefaultRequestHeaders.Add("Referer", "https://aurionweb.ensiie.fr/faces/Planning.xhtml");

            content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("javax.faces.partial.ajax", "true"),
                new KeyValuePair<string, string>("javax.faces.source", calendarForm),
                new KeyValuePair<string, string>("javax.faces.partial.execute", calendarForm),
                new KeyValuePair<string, string>(calendarForm, calendarForm),
                new KeyValuePair<string, string>("javax.faces.partial.render", calendarForm),
                new KeyValuePair<string, string>("form", "form"),
                new KeyValuePair<string, string>(calendarForm + "_view", "month"),
                new KeyValuePair<string, string>("form:offsetFuseauNavigateur", "-7200000"),
                new KeyValuePair<string, string>("javax.faces.ViewState", viewState),
                new KeyValuePair<string, string>(calendarForm + "_end", "3155760000000"),
                new KeyValuePair<string, string>(calendarForm + "_start", "0"),
                new KeyValuePair<string, string>("form:onglets_activeIndex", "0"),
                new KeyValuePair<string, string>("form:onglets_scrollState", "0"),
                new KeyValuePair<string, string>("form:largeurDivCenter", "1620")
            });
            response = Client.PostAsync("/faces/Planning.xhtml", content).Result;
            response.EnsureSuccessStatusCode();
            responseBody = await response.Content.ReadAsStringAsync();
            return (responseBody, name, calendarForm);
        }

        private static string ParseExecution(string body)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            return doc.DocumentNode.Descendants().FirstOrDefault(x => x.Attributes["name"]?.Value == "execution")
                ?.Attributes["value"]?.Value;
        }

        private static string ParseViewState(string body)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            return doc.DocumentNode.Descendants().FirstOrDefault(x => x.Id == "j_id1:javax.faces.ViewState:0")
                ?.Attributes["value"].Value;
        }

        private static string ParseMainMenuForm(string body)
        {
            return Regex.Match(body,
                    @"chargerSousMenu = function\(\) {PrimeFaces.ab\({s:""(form:j_idt\d+?)"",f:""form""").Groups[1]
                .Value;
        }

        private static string ParseCalendarMenuForm(string body)
        {
            return Regex.Match(body,
                    @"<li class=""ui-widget ui-menuitem ui-corner-all ui-menu-parent (submenu_\d+?)"" role=""menuitem"" aria-haspopup=""true""><a href=""#"" class=""ui-menuitem-link ui-submenu-link ui-corner-all"" tabindex=""-1""><span class=""ui-menuitem-text"">Emploi du temps<\/span>")
                .Groups[1]
                .Value;
        }

        private static string ParseCalendarSubMenuId(string body)
        {
            return Regex.Match(body,
                    @"PrimeFaces\.addSubmitParam\('form',{'form:sidebar':'form:sidebar','form:sidebar_menuid':'(\d_\d)'}\)\.submit\('form'\);return false;""><span class=""ui-menuitem-text"">Planning d'un étudiant au choix<\/span>")
                .Groups[1]
                .Value;
        }

        private static string ParseSelectCalendarForm(string body)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            return doc.DocumentNode.Descendants().FirstOrDefault(x =>
                x.Attributes["class"]?.Value == "ui-datatable ui-widget tableFavoriClass  ui-datatable-reflow")?.Id;
        }

        private static string ParseSelectCalendarButtonForm(string body)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            return doc.DocumentNode.Descendants().FirstOrDefault(x =>
                x.Attributes["class"]?.Value ==
                "ui-button ui-widget ui-state-default ui-corner-all ui-button-text-icon-left GreenButton")?.Id;
        }

        private static string ParseCalendarForm(string body)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            return doc.DocumentNode.Descendants().FirstOrDefault(x => x.Attributes["class"]?.Value == "schedule")?.Id;
        }

        private static string ParseName(string body)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            return doc.DocumentNode.Descendants().FirstOrDefault(x => x.Attributes["class"]?.Value == "divCurrentUser")
                ?.InnerText;
        }

        private static string ParseJsonBody(string body, string calendarForm)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(body);
            var data = doc.DocumentNode.Descendants("update").FirstOrDefault(x => x.Id == calendarForm)?.InnerHtml;

            return Regex.Match(data ?? throw new InvalidOperationException(), @"^<!\[CDATA\[\{""events"" : (.+?)\}]]>$")
                .Groups[1].Value;
        }

        private static List<AurionCalendarEvent> DeserializeCalendar(string jsonBody)
        {
            var listCalendarEvent = (List<AurionCalendarEvent>) JsonSerializer.Deserialize(jsonBody,
                typeof(List<AurionCalendarEvent>), new JsonSerializerOptions
                {
                    IgnoreNullValues = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
            return listCalendarEvent;
        }

        private static void ExportToIcs(List<AurionCalendarEvent> listCalendarEvent, string path)
        {
            var calendar = new Calendar();
            calendar.AddTimeZone(new VTimeZone("Europe/Paris"));
            foreach (var calendarEvent in listCalendarEvent)
            {
                var evt = new CalendarEvent
                {
                    Start = new CalDateTime(Convert.ToDateTime(calendarEvent.Start)),
                    End = new CalDateTime(Convert.ToDateTime(calendarEvent.End)),
                    IsAllDay = calendarEvent.AllDay,
                    Summary = calendarEvent.Title
                };
                calendar.Events.Add(evt);
            }

            var serializer = new CalendarSerializer();
            var serializedCalendar = serializer.SerializeToString(calendar);

            using (var file = new StreamWriter(path))
            {
                file.Write(serializedCalendar);
            }
        }
    }
}