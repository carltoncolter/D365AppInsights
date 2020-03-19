namespace JLattimer.D365AppInsights
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Globalization;
    using System.IO;
    using System.Net.Http;
    using System.Text;
    using System.Xml;

    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    public class AiLogger
    {
        private static HttpClient _httpClient;

        private string _authenticatedUserId;

        private bool _disableDependencyTracking;

        private bool _disableEventTracking;

        private bool _disableExceptionTracking;

        private bool _disableMetricTracking;

        private bool _disableTraceTracking;

        private bool _enableDebug;

        private string _instrumentationKey;

        private string _loggingEndpoint;

        private int _percentLoggedDependency;

        private int _percentLoggedEvent;

        private int _percentLoggedException;

        private int _percentLoggedMetric;

        private int _percentLoggedTrace;

        private ITracingService _tracingService;

        /// <summary>
        ///     Initializes a new instance of the <see cref="AiLogger" /> class.
        /// </summary>
        /// <param name="aiSetupJson">AiSetup json.</param>
        /// <param name="service">D365 IOrganizationService.</param>
        /// <param name="tracingService">D365 ITracingService.</param>
        /// <param name="executionContext">D365 IExecutionContext (IPluginExecutionContext or IWorkflowContext).</param>
        /// <param name="pluginStage">Plug-in stage from context</param>
        /// <param name="workflowCategory">Workflow category from context</param>
        public AiLogger(
            string aiSetupJson,
            IOrganizationService service,
            ITracingService tracingService,
            IExecutionContext executionContext,
            int? pluginStage,
            int? workflowCategory)
        {
            ValidateContextSpecific(pluginStage, workflowCategory);
            var aiConfig = new AiConfig(aiSetupJson);
            this.SetupAiLogger(aiConfig, service, tracingService, executionContext, pluginStage, workflowCategory);
        }

        /// <summary>
        ///     Initializes a new instance of the <see cref="AiLogger" /> class.
        /// </summary>
        /// <param name="aiConfig">AiConfiguration.</param>
        /// <param name="service">D365 IOrganizationService.</param>
        /// <param name="tracingService">D365 ITracingService.</param>
        /// <param name="executionContext">D365 IExecutionContext (IPluginExecutionContext or IWorkflowContext).</param>
        /// <param name="pluginStage">Plug-in stage from context</param>
        /// <param name="workflowCategory">Workflow category from context</param>
        public AiLogger(
            AiConfig aiConfig,
            IOrganizationService service,
            ITracingService tracingService,
            IExecutionContext executionContext,
            int? pluginStage,
            int? workflowCategory)
        {
            ValidateContextSpecific(pluginStage, workflowCategory);
            this.SetupAiLogger(aiConfig, service, tracingService, executionContext, pluginStage, workflowCategory);
        }

        public AiProperties EventProperties { get; private set; }

        public static void AppendValue(IOrganizationService service, object typeFullname, StringBuilder sb, object value)
        {
            switch (typeFullname)
            {
                case "System.Decimal":
                case "System.Double":
                case "System.Int32":
                case "System.Boolean":
                case "System.Single":
                case "System.String":
                case "System.Guid":
                    sb.Append('"');
                    sb.Append(value.ToString().Replace("\"", "&quot;"));
                    sb.Append('"');
                    break;
                case "System.DateTime":
                    sb.Append('"');
                    sb.Append(((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss \"GMT\"zzz"));
                    sb.Append('"');
                    break;
                case "Microsoft.Xrm.Sdk.EntityReference":
                    var entityReference = (EntityReference)value;
                    sb.Append($"Id: \"{entityReference.Id}\", LogicalName: \"{entityReference.LogicalName}\"");
                    break;
                case "Microsoft.Xrm.Sdk.Entity":
                    var entity = value as Entity;
                    if (entity == null)
                    {
                        sb.Append("NULL");
                    }
                    else
                    {
                        var tab = "     ";
                        sb.Append("\r\n{\r\n");
                        sb.Append($"{tab}Id: {entity.Id},\r\n{tab}LogicalName: {entity.LogicalName}");
                        foreach (var attr in entity.Attributes)
                        {
                            var attrType = (object)attr.Value?.GetType().FullName;
                            sb.Append($",\r\n{tab}{attr.Key}: ");
                            AppendValue(service, attrType, sb, attr.Value);
                        }

                        sb.AppendLine("\r\n}");
                    }

                    break;
                case "Microsoft.Xrm.Sdk.Query.ColumnSet":
                    sb.Append("[");
                    var colSet = value as ColumnSet;
                    if (colSet?.Columns.Count > 0)
                    {
                        foreach (var col in colSet.Columns) sb.Append($"\"{col}\", ");

                        sb.Remove(sb.Length - 2, 2);
                    }

                    sb.Append(" ]");
                    break;
                case "Microsoft.Xrm.Sdk.Query.QueryExpression":
                    QueryExpressionToFetchXmlRequest request = new QueryExpressionToFetchXmlRequest();
                    request.Query = (QueryExpression)value;
                    
                    var queryExpression = (QueryExpression)value; 
                    var top = queryExpression.TopCount.HasValue ? queryExpression.TopCount.ToString() : null;
                    sb.Append($"EntityName: \"{queryExpression.EntityName}\", Top: \"{top}\"");

                    sb.Append(", Columns:");
                    AppendValue(service, "Microsoft.Xrm.Sdk.Query.ColumnSet", sb, queryExpression.ColumnSet);
                    sb.AppendLine(", ");
                    var response = (QueryExpressionToFetchXmlResponse) service.Execute(request);
                    sb.Append("Query: \r\n");
                    sb.Append(FormatXml(response.FetchXml));
                    break;
                case "Microsoft.Xrm.Sdk.Query.FetchExpression":
                    var fetchExpression = (FetchExpression)value;
                    sb.Append("Query: \r\n");
                    sb.Append(FormatXml(fetchExpression.Query));
                    break;
                case "Microsoft.Xrm.Sdk.EntityCollection":
                    var entityCollection = (EntityCollection)value;
                    sb.Append("[");
                    foreach (var e in entityCollection.Entities)
                    {
                        sb.AppendLine(string.Empty);
                        AppendValue(service, e.GetType().FullName, sb, e);
                        sb.Append(", ");
                    }

                    if (entityCollection.Entities.Count > 0) sb.Remove(sb.Length - 2, 2);

                    break;
                case "Microsoft.Xrm.Sdk.OptionSetValue":
                    sb.Append('"');
                    sb.Append(((OptionSetValue)value).Value);
                    sb.Append('"');
                    break;
                case "Microsoft.Xrm.Sdk.Money":
                    sb.Append('"');
                    sb.Append(((Money)value).Value.ToString(CultureInfo.CurrentCulture));
                    sb.Append('"');
                    break;
                case "Microsoft.Xrm.Sdk.ConcurrencyBehavior":
                    sb.Append('"');
                    sb.Append(((ConcurrencyBehavior)value).ToString());
                    sb.Append('"');
                    break;
                default:
                    sb.Append($"\"Undefined Type - {typeFullname}, Value: {value}\"");
                    break;
            }
        }

        /// <summary>
        ///     Writes a dependency message to Application Insights.
        /// </summary>
        /// <param name="name">The dependency name or absolute URL.</param>
        /// <param name="method">The HTTP method (only logged with URL).</param>
        /// <param name="type">The type of dependency (Ajax, HTTP, SQL, etc.).</param>
        /// <param name="duration">The duration in ms of the dependent event.</param>
        /// <param name="resultCode">The result code, HTTP or otherwise.</param>
        /// <param name="success">Set to <c>true</c> if the dependent event was successful, <c>false</c> otherwise.</param>
        /// <param name="data">Any other data associated with the dependent event.</param>
        /// <param name="timestamp">The UTC timestamp of the dependent event (default = DateTime.UtcNow).</param>
        /// <returns><c>true</c> if successfully logged, <c>false</c> otherwise.</returns>
        public bool WriteDependency(
            string name,
            string method,
            string type,
            int duration,
            int? resultCode,
            bool success,
            string data,
            DateTime? timestamp = null)
        {
            if (!this.Log("Dependency", this._disableDependencyTracking, this._percentLoggedDependency))
                return true;

            timestamp = timestamp ?? DateTime.UtcNow;

            var dependency = new AiDependency(
                this.EventProperties,
                name,
                method,
                type,
                duration,
                resultCode,
                success,
                data);

            var json = this.GetDependencyJsonString(timestamp.Value, dependency, null);

            if (this._enableDebug)
                this._tracingService.Trace($"DEBUG: Application Insights JSON: {CreateJsonDataLog(json)}");

            return this.SendToAi(json);
        }

        /// <summary>
        ///     Writes an event message to Application Insights.
        /// </summary>
        /// <param name="name">The event name.</param>
        /// <param name="measurements">The associated measurements.</param>
        /// <param name="timestamp">The UTC timestamp of the event (default = DateTime.UtcNow).</param>
        /// <returns><c>true</c> if successfully logged, <c>false</c> otherwise.</returns>
        public bool WriteEvent(
            string name,
            Dictionary<string, double> measurements,
            DateTime? timestamp = null,
            string operationName = null,
            string operationId = null)
        {
            if (!this.Log("Event", this._disableEventTracking, this._percentLoggedEvent))
                return true;

            timestamp = timestamp ?? DateTime.UtcNow;

            var aiEvent = new AiEvent(this.EventProperties, name);

            var json = this.GetEventJsonString(timestamp.Value, aiEvent, measurements, operationName, operationId);

            if (this._enableDebug)
                this._tracingService.Trace($"DEBUG: Application Insights JSON: {CreateJsonDataLog(json)}");

            return this.SendToAi(json);
        }

        /// <summary>
        ///     Writes exception data to Application Insights.
        /// </summary>
        /// <param name="exception">The exception being logged.</param>
        /// <param name="aiExceptionSeverity">The severity level <see cref="AiExceptionSeverity" />.</param>
        /// <param name="timestamp">The UTC timestamp of the event (default = DateTime.UtcNow).</param>
        /// <returns><c>true</c> if successfully logged, <c>false</c> otherwise.</returns>
        public bool WriteException(
            Exception exception,
            AiExceptionSeverity aiExceptionSeverity,
            DateTime? timestamp = null)
        {
            if (!this.Log("Exception", this._disableExceptionTracking, this._percentLoggedException))
                return true;

            timestamp = timestamp ?? DateTime.UtcNow;

            var aiException = new AiException(exception, aiExceptionSeverity);

            var json = this.GetExceptionJsonString(timestamp.Value, aiException);

            if (this._enableDebug)
                this._tracingService.Trace($"DEBUG: Application Insights JSON: {CreateJsonDataLog(json)}");

            return this.SendToAi(json);
        }

        /// <summary>
        ///     Writes a metric message to Application Insights.
        /// </summary>
        /// <param name="name">The metric name.</param>
        /// <param name="value">The metric value.</param>
        /// <param name="count">The count of metrics being logged (default = 1).</param>
        /// <param name="min">The minimum value of metrics being logged (default = value).</param>
        /// <param name="max">The maximum value of metrics being logged (default = value).</param>
        /// <param name="stdDev">The standard deviation of metrics being logged (default = 0).</param>
        /// <param name="timestamp">The UTC timestamp of the event (default = DateTime.UtcNow).</param>
        /// <returns><c>true</c> if successfully logged, <c>false</c> otherwise.</returns>
        public bool WriteMetric(
            string name,
            double value,
            double? count = null,
            double? min = null,
            double? max = null,
            double? stdDev = null,
            DateTime? timestamp = null)
        {
            if (!this.Log("Metric", this._disableMetricTracking, this._percentLoggedMetric))
                return true;

            timestamp = timestamp ?? DateTime.UtcNow;

            var metric = new AiMetric(name, value, count, min, max, stdDev);

            var json = this.GetMetricJsonString(timestamp.Value, metric);

            if (this._enableDebug)
                this._tracingService.Trace($"DEBUG: Application Insights JSON: {CreateJsonDataLog(json)}");

            return this.SendToAi(json);
        }

        /// <summary>
        ///     Writes a trace message to Application Insights.
        /// </summary>
        /// <param name="message">The trace message.</param>
        /// <param name="aiTraceSeverity">The severity level <see cref="AiTraceSeverity" />.</param>
        /// <param name="timestamp">The UTC timestamp of the event (default = DateTime.UtcNow).</param>
        /// <returns><c>true</c> if successfully logged, <c>false</c> otherwise.</returns>
        public bool WriteTrace(string message, AiTraceSeverity aiTraceSeverity, DateTime? timestamp = null)
        {
            if (!this.Log("Trace", this._disableTraceTracking, this._percentLoggedTrace))
                return true;

            timestamp = timestamp ?? DateTime.UtcNow;

            var aiTrace = new AiTrace(this.EventProperties, message, aiTraceSeverity);

            var json = this.GetTraceJsonString(timestamp.Value, aiTrace);

            if (this._enableDebug)
                this._tracingService.Trace($"DEBUG: Application Insights JSON: {CreateJsonDataLog(json)}");

            return this.SendToAi(json);
        }

        private static string CreateJsonDataLog(string json)
        {
            return json.Replace("{", "{{").Replace("}", "}}");
        }

        private static string FormatXml(string xml)
        {
            var result = string.Empty;

            var mStream = new MemoryStream();
            var writer = new XmlTextWriter(mStream, Encoding.Unicode);
            var document = new XmlDocument();

            try
            {
                // Load the XmlDocument with the XML.
                document.LoadXml(xml);

                writer.Formatting = Formatting.Indented;

                // Write the XML into a formatting XmlTextWriter
                document.WriteContentTo(writer);
                writer.Flush();
                mStream.Flush();

                // Have to rewind the MemoryStream in order to read
                // its contents.
                mStream.Position = 0;

                // Read MemoryStream contents into a StreamReader.
                var sReader = new StreamReader(mStream);

                // Extract the text from the StreamReader.
                var formattedXml = sReader.ReadToEnd();

                result = formattedXml;
            }
            catch (XmlException)
            {
                // Handle the exception
            }

            mStream.Close();
            writer.Close();

            return result;
        }

        private static string GetVersion(IOrganizationService service)
        {
            var request = new RetrieveVersionRequest();
            var response = (RetrieveVersionResponse)service.Execute(request);

            return response.Version;
        }

        private static string GetWorkflowCategoryName(int category)
        {
            switch (category)
            {
                case 0:
                    return "Workflow";
                case 1:
                    return "Dialog";
                case 2:
                    return "Business Rule";
                case 3:
                    return "Action";
                case 4:
                    return "Business Process Flow";
                default:
                    return "Unknown";
            }
        }

        private static bool InLogThreshold(int threshold)
        {
            if (threshold == 100)
                return true;
            if (threshold == 0)
                return false;

            var random = new Random();
            var number = random.Next(1, 100);

            return number <= threshold;
        }

        private static string InsertMetricsJson(Dictionary<string, double> measurements, string json)
        {
            if (measurements == null)
                return json;

            var replacement = "\"measurements\":{";

            var i = 0;
            foreach (var keyValuePair in measurements)
            {
                i++;
                replacement += $"\"{keyValuePair.Key}\": {keyValuePair.Value}";
                if (i != measurements.Count)
                    replacement += ", ";
            }

            replacement += "}";

            json = json.Replace("\"measurements\":null", replacement);

            return json;
        }

        private static string TrimPropertyValueLength(string value)
        {
            return value.Length > 8192 ? value.Substring(0, 8191) : value;
        }

        [SuppressMessage("ReSharper", "ParameterOnlyUsedForPreconditionCheck.Local")]
        private static void ValidateContextSpecific(int? pluginStage, int? workflowCategory)
        {
            if (pluginStage == null && workflowCategory == null)
                throw new InvalidPluginExecutionException(
                    "Either Plug-in Stage or Workflow Category must be passed to AiLogger");
        }

        private void AddExecutionContextDetails(
            IOrganizationService service,
            IExecutionContext executionContext,
            int? pluginStage,
            int? workflowCategory)
        {
            this.EventProperties.ImpersonatingUserId = executionContext.UserId.ToString();
            this.EventProperties.CorrelationId = executionContext.CorrelationId.ToString();
            this.EventProperties.Message = executionContext.MessageName;
            this.EventProperties.Mode = AiProperties.GetModeName(executionContext.Mode);
            this.EventProperties.Depth = executionContext.Depth;
            this.EventProperties.InputParameters = this.TraceParameters(true, service, executionContext);
            this.EventProperties.OutputParameters = this.TraceParameters(false, service, executionContext);
            this.EventProperties.OperationId = executionContext.OperationId.ToString();
            this.EventProperties.OperationCreatedOn = executionContext.OperationCreatedOn.ToUniversalTime()
                .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'");
            this.EventProperties.OrganizationId = executionContext.OrganizationId.ToString();

            if (pluginStage != null)
            {
                var stage = pluginStage.Value;
                this.AddPluginExecutionContextDetails(stage, executionContext as IPluginExecutionContext);
            }

            if (workflowCategory != null)
            {
                var category = workflowCategory.Value;
                this.AddWorkflowExecutionContextDetails(category);
            }
        }

        private void AddPluginExecutionContextDetails(int stage, IPluginExecutionContext executionContext)
        {
            this.EventProperties.Source = "Plug-in";
            this.EventProperties.Stage = AiProperties.GetStageName(stage);
            if (executionContext == null) return;

            this.EventProperties.ParentCorrelationId =
                (executionContext?.ParentContext?.CorrelationId ?? executionContext?.CorrelationId).ToString();

            // this.AddPluginExecutionHistory(executionContext);
        }

        private void AddPluginExecutionHistory(IPluginExecutionContext executionContext)
        {
            // This doesn't work... need to figure out something that can work here.
            DateTime? end = null;
            if (!executionContext.SharedVariables.ContainsKey("StartedOn"))
                executionContext.SharedVariables.Add("StartedOn", DateTime.UtcNow);
            else end = (DateTime?)executionContext.SharedVariables["StartedOn"];
            var past = new Stack<AiPluginProperties>();
            if (this.EventProperties.Depth > 1)
            {
                var currentContext = executionContext;
                var last = currentContext;
                while (currentContext != null)
                {
                    past.Push(
                        new AiPluginProperties
                            {
                                OperationId = currentContext.OperationId.ToString(),
                                ImpersonatingUserId = currentContext.UserId.ToString(),
                                CorrelationId = currentContext.CorrelationId.ToString(),
                                Message = currentContext.MessageName,
                                Mode = AiProperties.GetModeName(currentContext.Mode),
                                Depth = currentContext.Depth,
                                EntityId = currentContext.PrimaryEntityId.ToString(),
                                EntityName = currentContext.PrimaryEntityName,
                                Stage = AiProperties.GetStageName(currentContext.Stage),
                                ParentCorrelationId = currentContext.ParentContext?.CorrelationId.ToString(),
                                OperationCreatedOn = currentContext.OperationCreatedOn.ToUniversalTime()
                                    .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'")
                            });
                    last = currentContext;
                    currentContext = currentContext.ParentContext;
                }

                var span = DateTime.UtcNow - (end ?? last.OperationCreatedOn.ToUniversalTime());

                this.WriteMetric("BackendDuration", span.TotalMilliseconds);

                this.EventProperties.History = past.ToArray();
                this.EventProperties.Duration = span.TotalMilliseconds.ToString();
            }
        }

        private void AddPrimaryPropertyValues(IExecutionContext executionContext, IOrganizationService service)
        {
            this.EventProperties.EntityName = executionContext.PrimaryEntityName;
            this.EventProperties.EntityId = executionContext.PrimaryEntityId.ToString();
            this.EventProperties.OrgName = executionContext.OrganizationName;
            this.EventProperties.OrgVersion = GetVersion(service);
        }

        private void AddWorkflowExecutionContextDetails(int workflowCategory)
        {
            this.EventProperties.Source = "Workflow";
            this.EventProperties.WorkflowCategory = GetWorkflowCategoryName(workflowCategory);
        }

        private string GetDependencyJsonString(DateTime timestamp, AiBaseData aiDependency, string operationName)
        {
            var logRequest = new AiLogRequest
                                 {
                                     Name =
                                         $"Microsoft.ApplicationInsights.{this._instrumentationKey}.RemoteDependency",
                                     Time = timestamp.ToString("O"),
                                     InstrumentationKey = this._instrumentationKey,
                                     Tags = new AiTags
                                                {
                                                    RoleInstance = string.Empty,
                                                    OperationName = operationName,
                                                    AuthenticatedUserId = this._authenticatedUserId
                                                },
                                     Data = new AiData { BaseType = "RemoteDependencyData", BaseData = aiDependency }
                                 };

            var json = SerializationHelper.SerializeObject<AiLogRequest>(logRequest);
            return json;
        }

        private string GetEventJsonString(
            DateTime timestamp,
            AiEvent aiEvent,
            Dictionary<string, double> measurements,
            string opName,
            string opid)
        {
            var logRequest = new AiLogRequest
                                 {
                                     Name = $"Microsoft.ApplicationInsights.{this._instrumentationKey}.Event",
                                     Time = timestamp.ToString("O"),
                                     InstrumentationKey = this._instrumentationKey,
                                     Tags = new AiTags
                                                {
                                                    RoleInstance = null,
                                                    OperationName = opName,
                                                    OperationId = opid,
                                                    AuthenticatedUserId = this._authenticatedUserId
                                                },
                                     Data = new AiData { BaseType = "EventData", BaseData = aiEvent }
                                 };

            var json = SerializationHelper.SerializeObject<AiLogRequest>(logRequest);
            json = InsertMetricsJson(measurements, json);

            return json;
        }

        private string GetExceptionJsonString(DateTime timestamp, AiException aiException)
        {
            var logRequest = new AiLogRequest
                                 {
                                     Name = $"Microsoft.ApplicationInsights.{this._instrumentationKey}.Exception",
                                     Time = timestamp.ToString("O"),
                                     InstrumentationKey = this._instrumentationKey,
                                     Tags = new AiTags
                                                {
                                                    RoleInstance = null,
                                                    OperationName = null,
                                                    AuthenticatedUserId = this._authenticatedUserId
                                                },
                                     Data = new AiData
                                                {
                                                    BaseType = "ExceptionData",
                                                    BaseData = new AiBaseData
                                                                   {
                                                                       Properties = this.EventProperties,
                                                                       Exceptions =
                                                                           new List<AiException> { aiException }
                                                                   }
                                                }
                                 };

            var json = SerializationHelper.SerializeObject<AiLogRequest>(logRequest);
            return json;
        }

        private string GetMetricJsonString(DateTime timestamp, AiMetric aiMetric)
        {
            var logRequest = new AiLogRequest
                                 {
                                     Name = $"Microsoft.ApplicationInsights.{this._instrumentationKey}.Metric",
                                     Time = timestamp.ToString("O"),
                                     InstrumentationKey = this._instrumentationKey,
                                     Tags = new AiTags
                                                {
                                                    RoleInstance = null,
                                                    OperationName = "xrmEvent",
                                                    AuthenticatedUserId = this._authenticatedUserId
                                                },
                                     Data = new AiData
                                                {
                                                    BaseType = "MetricData",
                                                    BaseData = new AiBaseData
                                                                   {
                                                                       Metrics = new List<AiMetric> { aiMetric },
                                                                       Properties = this.EventProperties
                                                                   }
                                                }
                                 };

            logRequest.Tags.OperationId = this.EventProperties.OperationId;

            var json = SerializationHelper.SerializeObject<AiLogRequest>(logRequest);
            return json;
        }

        private string GetTraceJsonString(DateTime timestamp, AiTrace aiTrace)
        {
            var logRequest = new AiLogRequest
                                 {
                                     Name = $"Microsoft.ApplicationInsights.{this._instrumentationKey}.Message",
                                     Time = timestamp.ToString("O"),
                                     InstrumentationKey = this._instrumentationKey,
                                     Tags = new AiTags
                                                {
                                                    OperationName = null,
                                                    RoleInstance = null,
                                                    AuthenticatedUserId = this._authenticatedUserId
                                                },
                                     Data = new AiData { BaseType = "MessageData", BaseData = aiTrace }
                                 };

            var json = SerializationHelper.SerializeObject<AiLogRequest>(logRequest);
            return json;
        }

        private bool Log(string type, bool disable, int threshold)
        {
            if (disable)
            {
                if (this._enableDebug)
                    this._tracingService.Trace($"DEBUG: Application Insights {type} not written: Disabled");
                return false;
            }

            var shouldLog = InLogThreshold(threshold);
            if (!shouldLog)
            {
                if (this._enableDebug)
                    this._tracingService.Trace(
                        $"DEBUG: Application Insights {type} not written: Threshold%: {threshold}");
                return false;
            }

            return true;
        }

        private bool SendToAi(string json)
        {
            try
            {
                HttpContent content = new StringContent(json, Encoding.UTF8, "application/x-json-stream");
                var response = _httpClient.PostAsync(this._loggingEndpoint, content).Result;

                if (response.IsSuccessStatusCode)
                    return true;

                this._tracingService?.Trace(
                    $"ERROR: Unable to write to Application Insights with response: {response.StatusCode.ToString()}: {response.ReasonPhrase}: Message: {CreateJsonDataLog(json)}");
                return false;
            }
            catch (Exception e)
            {
                this._tracingService?.Trace(CreateJsonDataLog(json), e);
                return false;
            }
        }

        private void SetupAiLogger(
            AiConfig aiConfig,
            IOrganizationService service,
            ITracingService tracingService,
            IExecutionContext executionContext,
            int? pluginStage,
            int? workflowCategory)
        {
            this._instrumentationKey = aiConfig.InstrumentationKey;
            this._loggingEndpoint = aiConfig.AiEndpoint;
            this._disableTraceTracking = aiConfig.DisableTraceTracking;
            this._disableExceptionTracking = aiConfig.DisableExceptionTracking;
            this._disableDependencyTracking = aiConfig.DisableDependencyTracking;
            this._disableEventTracking = aiConfig.DisableEventTracking;
            this._disableMetricTracking = aiConfig.DisableMetricTracking;
            this._enableDebug = aiConfig.EnableDebug;
            this._percentLoggedTrace = aiConfig.PercentLoggedTrace;
            this._percentLoggedMetric = aiConfig.PercentLoggedMetric;
            this._percentLoggedEvent = aiConfig.PercentLoggedEvent;
            this._percentLoggedException = aiConfig.PercentLoggedException;
            this._percentLoggedDependency = aiConfig.PercentLoggedDependency;
            var disableContextParameterTracking = aiConfig.DisableContextParameterTracking;
            this._authenticatedUserId = executionContext.InitiatingUserId.ToString();
            this._tracingService = tracingService;
            _httpClient = HttpHelper.GetHttpClient();

            this.EventProperties = new AiProperties();
            this.AddPrimaryPropertyValues(executionContext, service);

            if (!disableContextParameterTracking)
                this.AddExecutionContextDetails(service, executionContext, pluginStage, workflowCategory);
        }

        private string TraceParameters(bool input, IOrganizationService service, IExecutionContext executionContext)
        {
            var parameters = input ? executionContext.InputParameters : executionContext.OutputParameters;

            if (parameters == null || parameters.Count == 0)
                return null;

            try
            {
                var sb = new StringBuilder();

                var i = 1;
                foreach (var parameter in parameters)
                {
                    var typeFullname = (object)parameter.Value?.GetType().FullName;
                    var parameterType = input ? "Input" : "Output";
                    sb.Append($"{parameterType} Parameter({typeFullname}): {parameter.Key}: ");
                    AppendValue(service, typeFullname, sb, parameter.Value);

                    i++;
                    if (i <= parameters.Count)
                        sb.Append(Environment.NewLine);
                }

                var result = sb.ToString();
                return result.Length > 8192 ? result.Substring(0, 8191) : result;
            }
            catch (Exception e)
            {
                this._tracingService.Trace($"ERROR: Tracing parameters: {e.Message}");
                return null;
            }
        }
    }
}