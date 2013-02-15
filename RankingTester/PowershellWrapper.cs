using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Xml.Linq;
using Microsoft.Office.Server.Search.Administration;

namespace Chaholl.RankingTester
{
    internal class PowershellWrapper
    {
        public static RankingModel AddOrUpdateRankingModel(Guid rankingModelId, XDocument model, SearchServiceApplication ssa)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();

            runspace.Open();
            runspace.SessionStateProxy.SetVariable("ssa", ssa);
            runspace.SessionStateProxy.SetVariable("rankmodel", model.ToString());

            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Add-PSSnapin Microsoft.Sharepoint.PowerShell");
            pipeline.Commands.AddScript("New-SPEnterpriseSearchRankingModel –SearchApplication $ssa -RankingModelXML $rankmodel");

            Collection<PSObject> results = pipeline.Invoke();

            runspace.Close();

            if (results.Count == 1)
            {
                //Get the ranking mode to ensure it's available
                return GetRankingModel(rankingModelId, ssa);
            }
            //couldn't add. Try an update
            return SetRankingModel(rankingModelId, model, ssa);
        }

        public static RankingModel GetRankingModel(Guid rankingModelId, SearchServiceApplication ssa)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();

            runspace.Open();
            runspace.SessionStateProxy.SetVariable("ssa", ssa);
            runspace.SessionStateProxy.SetVariable("rankingModelId", rankingModelId.ToString());


            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Add-PSSnapin Microsoft.Sharepoint.PowerShell");
            pipeline.Commands.AddScript("Get-SPEnterpriseSearchRankingModel –SearchApplication $ssa -Identity $rankingModelId");

            Collection<PSObject> results = pipeline.Invoke();

            runspace.Close();

            if (results.Count == 1)
            {
                return results[0].BaseObject as RankingModel;
            }

            return null;
        }

        public static void RemoveRankingModel(Guid rankingModelId, SearchServiceApplication ssa)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();

            runspace.Open();
            runspace.SessionStateProxy.SetVariable("ssa", ssa);
            runspace.SessionStateProxy.SetVariable("rankingModelId", rankingModelId.ToString());


            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Add-PSSnapin Microsoft.Sharepoint.PowerShell");
            pipeline.Commands.AddScript(
                "Remove-SPEnterpriseSearchRankingModel –SearchApplication $ssa -Identity $rankingModelId -Confirm:$false");

            pipeline.Invoke();

            runspace.Close();
        }

        private static RankingModel SetRankingModel(Guid rankingModelId, XDocument model, SearchServiceApplication ssa)
        {
            Runspace runspace = RunspaceFactory.CreateRunspace();

            runspace.Open();
            runspace.SessionStateProxy.SetVariable("ssa", ssa);
            runspace.SessionStateProxy.SetVariable("rankmodel", model.ToString());
            runspace.SessionStateProxy.SetVariable("identity", rankingModelId.ToString());

            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Add-PSSnapin Microsoft.Sharepoint.PowerShell");
            pipeline.Commands.AddScript(
                "Set-SPEnterpriseSearchRankingModel –SearchApplication $ssa -Identity $identity -RankingModelXML $rankmodel");

            pipeline.Invoke();

            runspace.Close();

            return GetRankingModel(rankingModelId, ssa);
        }
    }
}