using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace fdx_PrimaryLeadUpdate
{
    public class fdx_PrimaryLeadUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            int step = 0;
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                Entity leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(context.UserId);

                    //convert field type to bool & check
                    bool isprimarylead = leadEntity.GetAttributeValue<bool>("fdx_primarylead");
                    EntityCollection contextLeadCollection;
                    if (isprimarylead == true)
                    {
                        step = 0;
                        //fetch groupid of context lead
                        QueryExpression contextLeadQuery = new QueryExpression();
                        contextLeadQuery.EntityName = "lead";
                        contextLeadQuery.ColumnSet = new ColumnSet("fdx_groupid");
                        contextLeadQuery.Criteria.AddCondition("leadid", ConditionOperator.Equal, leadEntity.Id);
                        contextLeadCollection = service.RetrieveMultiple(contextLeadQuery);
                        step = 1;
                        if (contextLeadCollection.Entities.Count > 0)
                        {
                            Entity contextLead = contextLeadCollection[0];

                            //fetch all leads with similar group id
                            QueryExpression connectedLeadsQuery = new QueryExpression();
                            connectedLeadsQuery.EntityName = "lead";
                            connectedLeadsQuery.ColumnSet = new ColumnSet("fdx_groupid", "fdx_primarylead");//add the group ID attribute***
                            connectedLeadsQuery.Criteria.AddFilter(LogicalOperator.And);
                            connectedLeadsQuery.Criteria.AddCondition("fdx_groupid", ConditionOperator.Equal, contextLead.Attributes["fdx_groupid"]);
                            connectedLeadsQuery.Criteria.AddCondition("leadid", ConditionOperator.NotEqual, contextLead.Id);
                            EntityCollection connectedLeadsCollection = service.RetrieveMultiple(connectedLeadsQuery);
                            step = 2;
                            if (connectedLeadsCollection.Entities.Count > 0)
                            {
                                //Loop through connected Leads
                                //for (int i = 0; i < connectedLeadsCollection.Entities.Count; i++)
                                foreach (Entity connectedLead in connectedLeadsCollection.Entities)
                                {
                                    //Entity connectedLead = connectedLeadsCollection[i];
                                    if (connectedLead.Attributes.Contains("fdx_primarylead") && (bool)connectedLead.Attributes["fdx_primarylead"] == true)
                                    {
                                        connectedLead.Attributes["fdx_primarylead"] = false;
                                        service.Update(connectedLead);
                                    }
                                    else if (!connectedLead.Attributes.Contains("fdx_primarylead"))
                                    {
                                        connectedLead.Attributes["fdx_primarylead"] = false;
                                        service.Update(connectedLead);
                                    }
                                }
                            }
                        }
                    }

                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred while setting primary lead. " + ex.Message);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("Lead_update: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}