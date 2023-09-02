using HtmlAgilityPack;
using System.Diagnostics;
using System.IO;

namespace CISABulletins
{
    public class Row
    {
        public string Description { get; set; }
        public string Published { get; set; }
        public string Score { get; set; }
        public string SourceAndPatch { get; set; }
    }
    public class Program
    {
        async static Task Main(string[] args)
        {
            string directory = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            directory = Path.Combine(directory, "CISABulletin");
            Directory.CreateDirectory(directory);

            int start, end;
            string innerHtml, table;
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "C# console program");
            var content = await client.GetStringAsync("https://www.cisa.gov/news-events/bulletins");
            string lastViewedFile = Path.Combine(directory, "CISA_LastViewed.txt");

            DateTime lastViewed = DateTime.MinValue;
            if (File.Exists(lastViewedFile))
            {
                DateTime.TryParse(File.ReadAllText(lastViewedFile), out lastViewed);
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(content);
            var nodes = doc.DocumentNode.SelectNodes("//article");
            var nodes2 = nodes[0].SelectNodes("//div/div/h3/a");
            string url = null;
            string dateText = string.Empty;
            if (lastViewed != DateTime.MinValue) 
            {
                DateTime fileDate;
                int index = 0;
                int i = nodes2.Count - 1;
                while(i >= 0)
                {
                    var text = nodes2[i].InnerText.Replace("\n", "").Trim();
                    dateText = text.Replace("Vulnerability Summary for the Week of ", "");
                    DateTime.TryParse(dateText, out fileDate);
                    if (fileDate > lastViewed)
                    {
                        index = i;
                        break;
                    }
                    i--;
                }
                url = "https://www.cisa.gov" + nodes2[index].Attributes[0].Value;
                File.WriteAllText(lastViewedFile, dateText);
            }
            else
            {
                var text = nodes2[nodes2.Count - 1].InnerText.Replace("\n", "").Trim();
                dateText = text.Replace("Vulnerability Summary for the Week of ", "");
                File.WriteAllText(lastViewedFile, dateText);
                url = "https://www.cisa.gov" + nodes2[nodes2.Count - 1].Attributes[0].Value;
            }

            client.DefaultRequestHeaders.Add("User-Agent", "C# console program");
            content = await client.GetStringAsync(url);

            doc = new HtmlDocument();
            doc.LoadHtml(content);
            var title = doc.DocumentNode.SelectSingleNode("//h1[@class='c-page-title__title']/span").InnerText;  // h2 = High Vulnerabilities

            nodes = doc.DocumentNode.SelectNodes("//div[@id='medium_v']");  // h2 = High Vulnerabilities
            Dictionary<string, List<Row>> highVulnerabilitiesTable = new Dictionary<string, List<Row>>();
            foreach (var node in nodes)
            {
                if (node.ChildNodes[0].Name == "h2")
                {
                    innerHtml = node.InnerHtml;
                    start = innerHtml.IndexOf("<table");
                    end = innerHtml.IndexOf("</table>");
                    table = innerHtml.Substring(start, end - start + 8);
                    highVulnerabilitiesTable = ParseTable(table);
                    break;
                }
            }

            nodes = doc.DocumentNode.SelectNodes("//div[@id='medium_v']");
            innerHtml = nodes[0].InnerHtml;
            start = innerHtml.IndexOf("<table");
            end = innerHtml.IndexOf("</table>");
            table = innerHtml.Substring(start, end - start + 8);
            Dictionary<string, List<Row>> mediumVulnerabilitiesTable = ParseTable(table);

            nodes = doc.DocumentNode.SelectNodes("//div[@id='low_v']");
            innerHtml = nodes[0].InnerHtml;
            start = innerHtml.IndexOf("<table");
            end = innerHtml.IndexOf("</table>");
            table = innerHtml.Substring(start, end - start + 8);
            Dictionary<string, List<Row>> lowVulnerabilitiesTable = ParseTable(table);

            nodes = doc.DocumentNode.SelectNodes("//div[@id='snya_v']");
            innerHtml = nodes[0].InnerHtml;
            start = innerHtml.IndexOf("<table");
            end = innerHtml.IndexOf("</table>");
            table = innerHtml.Substring(start, end - start + 8);
            Dictionary<string, List<Row>> notAssignedVulnerabilitiesTable = ParseTable(table);

            string filename = Path.Combine(directory, "CISABulletin.htm");

            using (StreamWriter writer = new StreamWriter(filename))
            {
                writer.WriteLine("<html>");
                writer.WriteLine("<head>");
                writer.WriteLine($"<h2><center>{title}</center></h2>");
                writer.WriteLine("<style>");
                writer.WriteLine("table { ");
                writer.WriteLine("\twidth: 90%; ");
                writer.WriteLine("\tborder-collapse: collapse; ");
                writer.WriteLine("\tmargin:50px auto;");
                writer.WriteLine("\t}");
                writer.WriteLine("");
                writer.WriteLine("th { ");
                writer.WriteLine("\tbackground: #3498db; ");
                writer.WriteLine("\tcolor: white; ");
                writer.WriteLine("\tfont-weight: bold; ");
                writer.WriteLine("\t}");
                writer.WriteLine("");
                writer.WriteLine("td, th { ");
                writer.WriteLine("\tpadding: 10px; ");
                writer.WriteLine("\tborder: 1px solid #ccc; ");
                writer.WriteLine("\ttext-align: left; ");
                writer.WriteLine("\tfont-size: 18px;");
                writer.WriteLine("\t}");
                writer.WriteLine("");
                writer.WriteLine(".labels tr td {");
                writer.WriteLine("\tbackground-color: #2cc16a;");
                writer.WriteLine("\tfont-weight: bold;");
                writer.WriteLine("\tcolor: #fff;");
                writer.WriteLine("}");
                writer.WriteLine("");
                writer.WriteLine(".label tr td label {");
                writer.WriteLine("\tdisplay: block;");
                writer.WriteLine("}");
                writer.WriteLine("");
                writer.WriteLine("");
                writer.WriteLine("[data-toggle=\"toggle\"] {");
                writer.WriteLine("\tdisplay: none;");
                writer.WriteLine("}");
                writer.WriteLine("</style>");
                writer.WriteLine("</head>");
                writer.WriteLine("<body>");
                WriteSection("High Vulnerabilities", highVulnerabilitiesTable, writer);
                writer.WriteLine("<br/>");
                WriteSection("Medium Vulnerabilities", mediumVulnerabilitiesTable, writer);
                writer.WriteLine("<br/>");
                WriteSection("Low Vulnerabilities", lowVulnerabilitiesTable, writer);
                writer.WriteLine("<br/>");
                WriteSection("Unassigned Vulnerabilities", notAssignedVulnerabilitiesTable, writer);
                writer.WriteLine("<br/>");
                writer.WriteLine("</body>");
                writer.WriteLine("</html>");
            }
            Process.Start("explorer", "\"" + filename + "\"");
        }

        public static void WriteSection(string sectionName, Dictionary<string, List<Row>> data, StreamWriter writer)
        {
            writer.WriteLine($"<h3><b>{sectionName}</b></h3>");
            var keys = data.Keys.OrderBy(x => x).ToList();
            foreach(string key in keys)
            {
                writer.WriteLine("<details>");
                writer.WriteLine($"<summary>{key}</summary>");
                writer.WriteLine("<table>");
                writer.WriteLine(" <thead>");
                writer.WriteLine("     <tr>");
                writer.WriteLine("         <th>Description</th>");
                writer.WriteLine("         <th>Published</th>");
                writer.WriteLine("         <th>CVSS Score</th>");
                writer.WriteLine("         <th>Source & Patch Info</th>");
                writer.WriteLine("     </tr>");
                writer.WriteLine(" </thead>");
                writer.WriteLine(" <tbody>");
                foreach (Row row in data[key])
                {
                    writer.WriteLine("     <tr>");
                    writer.WriteLine($"         <td>{row.Description}</td>");
                    writer.WriteLine($"         <td>{row.Published}</td>");
                    writer.WriteLine($"         <td>{row.Score}</td>");
                    writer.WriteLine($"         <td>{row.SourceAndPatch}</td>");
                    writer.WriteLine("     </tr>");
                }
                writer.WriteLine(" </tbody>");
                writer.WriteLine("</table>");
                writer.WriteLine("</details>");
            }
        }

        public static Dictionary<string, List<Row>> ParseTable(string html)
        {
            Dictionary<string, List<Row>> tableData = new Dictionary<string, List<Row>>();
            var doc = new HtmlDocument();
            doc.LoadHtml(html);
            var nodes = doc.DocumentNode.SelectNodes("//table/tbody/tr"); 
            foreach (var node in nodes)
            {
                string product = node.ChildNodes[0].InnerHtml;
                int index = product.IndexOf("<br>");
                if (index != -1)
                {
                    product = product.Substring(0, index);
                }
                if (!tableData.Keys.Contains(product))
                {
                    tableData.Add(product, new List<Row>());
                }
                Row row = new Row();
                row.Description = node.ChildNodes[1].InnerHtml;
                row.Published = node.ChildNodes[2].InnerHtml;
                row.Score = node.ChildNodes[3].InnerHtml;
                row.SourceAndPatch = node.ChildNodes[4].InnerHtml;
                tableData[product].Add(row);
            }
            return tableData;
        }
    }
}