using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tricentis.TCAddOns;
using Tricentis.TCAPIObjects.Objects;

namespace Tricentis.Goodies.Integration.ToscaDefectHelper
{
    class OpenDefectTask : TCAddOnTask
    {
        public override Type ApplicableType { get { return typeof(Issue); } }

        public override string Name
        {
            get
            {
                return "Open Defect";
            }
        }

        public override bool RequiresChangeRights
        {
            get
            {
                return false;
            }
        }

        public override bool IsTaskPossible(TCObject obj)
        {
            string issueId;

            try
            {
                issueId = obj.GetAttributeValue(Properties.Settings.Default.ExternalDefectIdFieldName);
            }
            catch (Exception)
            {
                issueId = null;
            }

            return !(string.IsNullOrWhiteSpace(issueId) || string.IsNullOrWhiteSpace(Properties.Settings.Default.OpenDefectCommand));
        }

        public override TCObject Execute(TCObject objectToExecuteOn, TCAddOnTaskContext taskContext)
        {
            string issueId = objectToExecuteOn.GetAttributeValue(Properties.Settings.Default.ExternalDefectIdFieldName);

            string openCmd = string.Format(Properties.Settings.Default.OpenDefectCommand, issueId);

            if (!string.IsNullOrWhiteSpace(openCmd))
            {
                ProcessStartInfo sInfo = new ProcessStartInfo(openCmd);
                Process.Start(sInfo); 
            }
            return null;
        }
    }
}
