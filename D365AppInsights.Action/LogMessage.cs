using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace D365AppInsights.Action
{
    using JLattimer.D365AppInsights;

    using Microsoft.Xrm.Sdk;

    public class LogMessage : PluginBase
    {
        #region Constructor/Configuration

        private readonly string _unsecureConfig;

        private readonly string _secureConfig;

        public LogMessage(string unsecure, string secure)
            : base(typeof(LogEvent))
        {
            _unsecureConfig = unsecure;
            _secureConfig = secure;
        }

        #endregion

        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            //if (localContext.PluginExecutionContext.MessageName != "Update") return;
            if (localContext.PluginExecutionContext.PrimaryEntityName == "plugintracelog"
                || localContext.PluginExecutionContext.PrimaryEntityName.Substring(0, 3) == "sdk")
                return; // abort plugin if not needed.
            try
            {
                if (localContext.PluginExecutionContext.MessageName.Equals("update", StringComparison.InvariantCultureIgnoreCase) &&
                    localContext.PluginExecutionContext.Stage == 10)
                {
                    if (!localContext.PluginExecutionContext.InputParameters.ContainsKey("ConcurrencyBehavior"))
                    {
                        localContext.PluginExecutionContext.InputParameters["ConcurrencyBehavior"] =
                            ConcurrencyBehavior.Default;
                    }
                }
                AiLogger aiLogger = new AiLogger(
                    _unsecureConfig,
                    localContext.OrganizationService,
                    localContext.TracingService,
                    localContext.PluginExecutionContext,
                    localContext.PluginExecutionContext.Stage,
                    null);

                var measurements =
                    new Dictionary<string, double> { { "Stage", localContext.PluginExecutionContext.Stage } };
                aiLogger.WriteEvent("XRM Message", measurements);
            }
            catch (Exception e)
            {
                localContext.TracingService.Trace($"Unhandled Exception: {e.Message}");
                //throw;
                //ActionHelpers.SetOutputParameters(localContext.PluginExecutionContext.OutputParameters, false, e.Message);
            }
        }
    }
}
