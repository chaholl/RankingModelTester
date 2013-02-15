using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Server.Search.Administration;
using Microsoft.SharePoint;
using Microsoft.Office.Server.Search.Query;

namespace Chaholl.RankingTester
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length <2)
            {
                ShowHelpText();
                return;
            }
           

            var server = args[0];
            var queryText = args[1];
            FileStream rankingModelFile = null;
            if (args.Length >= 3)
            {
                rankingModelFile = File.OpenRead(args[2]);
            }

            DoQuery(server, queryText, Guid.Empty);

            if (rankingModelFile != null)
            {
                var ssa =
               SearchService.Service.SearchApplications.OfType<SearchServiceApplication>()
                            .First(s => s.SearchApplicationType != SearchServiceApplicationType.ExtendedConnector && s.FASTAdminProxy == null);

                rankingModelFile.Seek(0, SeekOrigin.Begin);
                var model = XDocument.Load(new XmlTextReader(rankingModelFile));

                if (model.Root == null) throw new Exception("Invalid ranking model");
                var idAttribute = model.Root.Attribute("id");
                Guid rankingModelId;
                if (idAttribute == null)
                {
                    //add a new ID
                    rankingModelId = Guid.NewGuid();
                    model.Root.Add(new XAttribute("id", rankingModelId.ToString()));
                }
                else
                {
                    rankingModelId = new Guid(idAttribute.Value);
                }

                XNamespace ns = "http://schemas.microsoft.com/office/2009/rankingModel";
                foreach (var element in model.Descendants(ns + "queryDependentFeature"))
                {
                    HookupPropertyId(element, ssa);
                }

                foreach (var element in model.Descendants(ns + "queryIndependentFeature"))
                {
                    HookupPropertyId(element, ssa);
                }

                RankingModel rankingModel = PowershellWrapper.AddOrUpdateRankingModel(rankingModelId, model, ssa);
                rankingModelFile.Close();

                if (rankingModel != null)
                {
                    //Although we've added the ranking model it isn't immediately available to search so we wait a bit
                    Thread.Sleep(30000);
                    try
                    {
                        DoQuery(server, queryText, rankingModel.ID);
                    }
                    finally
                    {
                        PowershellWrapper.RemoveRankingModel(rankingModel.ID, ssa);
                    }
                }

            }
        }

        private static void ShowHelpText()
        {
            Console.WriteLine();
            Console.ForegroundColor=ConsoleColor.Yellow;

            using (var helpTextStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Chaholl.RankingTester.help.txt"))
            {
                if (helpTextStream != null)
                {
                    var reader = new StreamReader(helpTextStream);
                    Console.Write(reader.ReadToEnd());
                }
            }

            Console.ForegroundColor = ConsoleColor.White;
        }

        private static void HookupPropertyId(XElement element, SearchServiceApplication ssa)
        {
            var pid = element.Attribute("pid");
            var name = element.Attribute("name");

            if ((pid == null || pid.Value == "?") && name != null)
            {
                //lookup the proeprty name
                var managedProperty = ssa.GetManagedProperties().FirstOrDefault(mp => mp.Name.StartsWith(name.Value, StringComparison.InvariantCultureIgnoreCase) && mp.Name.Length == name.Value.Length);
                if (managedProperty != null)
                {
                    if (pid == null)
                    {
                        element.Add(new XAttribute("pid", managedProperty.Pid));
                    }
                    else
                    {
                        pid.Value = managedProperty.Pid.ToString(CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    throw new Exception("Invalid managed property name: " + name.Value);
                }
            }
        }

        private static void DoQuery(string server, string queryText, Guid rankingModelId)
        {
            //We've got time to kill so we'll re-get the SPSite object
            using (var site = new SPSite(server))
            {
                Console.ForegroundColor = ConsoleColor.White;
                var query = new KeywordQuery(site) { QueryText = queryText, ResultTypes = ResultType.RelevantResults };

                query.SelectProperties.Clear();
                query.SelectProperties.Add("RANK");
                query.SelectProperties.Add("Filename");
                query.SelectProperties.Add("Title");

                if (rankingModelId != Guid.Empty)
                {
                    query.RankingModelId = rankingModelId.ToString();
                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine();
                    Console.WriteLine("Query with ranking model {0}", rankingModelId);
                }

                ResultTableCollection resultTableCollection = query.Execute();
                ResultTable resultTable = resultTableCollection[ResultType.RelevantResults];

                Console.WriteLine();
                Console.WriteLine("Results:");
                Console.WriteLine();

                while (resultTable.Read())
                {
                    Console.WriteLine("Rank:{0,8}\tTitle:{1,30}\tFilename:{2,20}", resultTable["RANK"], resultTable["TITLE"].ToString().Truncate(30),
                                      resultTable["FILENAME"].ToString().Truncate(20));
                }

                Console.WriteLine("Returned {0} results", resultTable.RowCount);
                Console.ForegroundColor = ConsoleColor.White;
            }
        }
    }

    public static class StringExt
    {
        public static string Truncate(this string value, int maxLength)
        {
            return value.Length <= maxLength ? value : value.Substring(0, maxLength - 1) + "…";
        }
    }
}
