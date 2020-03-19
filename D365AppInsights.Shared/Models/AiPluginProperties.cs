using System;
using System.Runtime.Serialization;

namespace JLattimer.D365AppInsights
{
    using System.Runtime.Serialization;
    using System.Threading.Tasks;

    using Microsoft.Xrm.Sdk;

    [DataContract]
    public class AiPluginProperties
    {
        [DataMember(Name = "operationId")]
        public string OperationId { get; set; }
        [DataMember(Name = "operationCreatedOn")]
        public string OperationCreatedOn { get; set; }
        //Define additional properties as needed
        [DataMember(Name = "impersonatingUserId")]
        public string ImpersonatingUserId { get; set; }

        [DataMember(Name = "entityId")]
        public string EntityId { get; set; }

        [DataMember(Name = "entityName")]
        public string EntityName { get; set; }

        [DataMember(Name = "correlationId")]
        public string CorrelationId { get; set; }

        [DataMember(Name = "message")]
        public string Message { get; set; }

        [DataMember(Name = "stage")]
        public string Stage { get; set; }

        [DataMember(Name = "mode")]
        public string Mode { get; set; }

        [DataMember(Name = "depth")]
        public int Depth { get; set; }

        [DataMember(Name = "parentCorrelationId")]
        public string ParentCorrelationId { get; set; }
    }
}
