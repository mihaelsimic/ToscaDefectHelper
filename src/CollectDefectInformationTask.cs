using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Tricentis.TCAddOns;
using Tricentis.TCAPIObjects.Objects;

namespace Tricentis.Goodies.Integration.ToscaDefectHelper
{
    internal class CollectDefectInformationTask : TCAddOnTask
    {

        #region Attributes
        public override Type ApplicableType
        {
            get
            {
                return typeof(Issue);
            }
        }

        public override bool IsTaskPossible(TCObject obj)
        {

            if (obj.GetType() == typeof(Issue) && ((Issue)obj).Type == IssueType.Defect)
            {
                return true;
            }
            else
            {
                return base.IsTaskPossible(obj);
            }
        }

        public override bool RequiresChangeRights
        {
            get
            {
                return true;
            }
        }

        public override string Name
        {
            get
            {
                return "Collect Defect Information";
            }
        } 
        #endregion

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            Issue target = objectToExecuteOn as Issue;

            if (target == null)
            {
                return null;
            }

            // This one just holds the brief info about failure(s)
            StringBuilder failureInfo = new StringBuilder();

            // This one holds the detailed info from the engine
            StringBuilder verboseLogInfo = new StringBuilder();

            // Holds all paths to automatically taken screenshots to be added to the defect
            List<string> attachments = new List<string>();

            // Update the status bar for the user
            string statusMessagePrefix = "Collecting data";
            taskContext.ShowStatusInfo(statusMessagePrefix);

            // Loop through all referenced test case logs
            foreach (IssueLink linkedTestCaseLogs in target.Links)
            {
                ExecutionTestCaseLog log = linkedTestCaseLogs.ExecutionTestCaseLog;
                //target.Description = log.AggregatedDescription;
                if (log != null)
                {
                    // Update the status bar for the user
                    taskContext.ShowStatusInfo(String.Format("{0}: {1}", statusMessagePrefix, log.DisplayedName));

                    if (verboseLogInfo.Length > 0)
                    {
                        // Just add a separator line between test case logs
                        verboseLogInfo.AppendLine();
                    }

                    // Format the log information so it reads nicely
                    FormatTestCaseLogDetails(verboseLogInfo, log);

                    AddFailureInfo(attachments, failureInfo, log);
                    
                }
            }

            if (failureInfo.Length > 0)
            {
                taskContext.ShowStatusInfo(String.Format("Adding details"));

                try
                {
                    target.SetAttibuteValue("Description", failureInfo.ToString());
                }
                catch (Exception exc)
                {
                    taskContext.ShowErrorMessage(this.Name, exc.Message);
                }
            }

            if (verboseLogInfo.Length > 0)
            {
                taskContext.ShowStatusInfo(String.Format("Adding verbose details"));

                try
                {
                    target.SetAttibuteValue(Properties.Settings.Default.FailureDetailsFieldName, verboseLogInfo.ToString());
                }
                catch 
                {
                    // Ignore in case the custom attribute does not exist
                }
            }

            AddAtachmentsAsOwnedFiles(taskContext, target, attachments);

            return null;
        }

        #region Helpers

        private void AddAtachmentsAsOwnedFiles(TCAddOnTaskContext taskContext, Issue target, List<string> attachments)
        {
            HashSet<string> exisingAttachments = new HashSet<string>();
            foreach (TCObject o in target.Search("->SUBPARTS:OwnedFile"))
            {
                exisingAttachments.Add(o.DisplayedName);
            }

            foreach (string filePath in attachments)
            {
                taskContext.ShowStatusInfo(String.Format("Adding file: '{0}'", filePath));
                try
                {
                    if (!exisingAttachments.Contains(Path.GetFileName(filePath)))
                    {
                        target.AttachFile(filePath, "Embedded"); 
                    }
                }
                catch (Exception exc)
                {
                    taskContext.ShowErrorMessage(this.Name, string.Format("Could not attach file '{0}'{1}{2}", filePath, Environment.NewLine, exc.Message));
                }
            }
        }

        private static void FormatTestCaseLogDetails(StringBuilder logDetails, ExecutionTestCaseLog log)
        {
            logDetails.AppendLine(log.DisplayedName);
            logDetails.AppendFormat("Log Info: {0}", log.LogInfo);
            logDetails.AppendLine();
            logDetails.AppendLine(log.AggregatedDescription);
            logDetails.AppendLine("--- end test case log ---");
        }

        private static void AddFailureInfo(List<string> attachments, StringBuilder failureInfo,  ExecutionTestCaseLog tcLog)
        {
            
            foreach (TCObject item in tcLog.Search(@"=>SUBPARTS"))
            {
                ExecutionLogEntry entry = item as ExecutionLogEntry;
                if (entry != null)
                {
                    if (entry.Result == ExecutionResult.Failed || entry.Result == ExecutionResult.Error)
                    {
                        string logInfoValue = entry.LogInfo;
                        if (!string.IsNullOrWhiteSpace(logInfoValue))
                        {
                            string testStepValueName = entry.GetAttributeValue("Name");
                            failureInfo.AppendFormat("{0}:{1}", testStepValueName, Environment.NewLine);
                            failureInfo.AppendLine(logInfoValue);
                            failureInfo.AppendLine();
                        }
                    }

                    if (!string.IsNullOrEmpty(entry.Detail))
                    {
                        attachments.Add(entry.Detail); 
                    }
                }
            }
        }
        #endregion
    }
}
