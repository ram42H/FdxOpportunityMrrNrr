using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxOpportunityMrrNrr
{
    public class OpportunityProduct_Create : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            int step = 0;

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity OppProductEntity = (Entity)context.InputParameters["Target"];

                if (OppProductEntity.LogicalName != "opportunityproduct")
                    return;

                Entity opportunityEntity = new Entity();
                EntityCollection oppProductMRR = new EntityCollection();
                EntityCollection oppProductNRR = new EntityCollection();
                EntityCollection oppProductSetupNet = new EntityCollection();
                Guid opportunityId = Guid.Empty;

                try
                {
                    step = 1;
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    step = 2;
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    //Retrieve Opportunity Id from Opportunity Product....
                    step = 3;
                    if(OppProductEntity.Attributes.Contains("opportunityid"))
                        opportunityId = ((EntityReference)OppProductEntity.Attributes["opportunityid"]).Id;

                    step = 4;
                    if (opportunityId != Guid.Empty)
                    {
                        //Get Sum MRR....
                        step = 5;
                        oppProductMRR = CRMQueryExpression.getMRR(opportunityId, service);

                        //Get Sum NRR....
                        step = 6;
                        oppProductNRR = CRMQueryExpression.getNRR(opportunityId, service);

                        //Get Sum Set-up Fee....
                        oppProductSetupNet = CRMQueryExpression.getSetUpFee(opportunityId, service);
                        tracingService.Trace("SetUpFee:- " + oppProductSetupNet.Entities.Count);

                        //Update opportunity....
                        step = 7;
                        opportunityEntity = new Entity
                        {
                            Id = opportunityId,
                            LogicalName = "opportunity"
                        };

                        step = 8;
                        if (oppProductMRR.Entities.Count > 0)
                        {
                            step = 81;
                            opportunityEntity["fdx_totalmrr"] = ((AliasedValue)oppProductMRR.Entities[0].Attributes["MRR"]).Value; 
                        }

                        step = 9;
                        if (oppProductNRR.Entities.Count > 0)
                        {
                            step = 91;
                            opportunityEntity["fdx_totalnrr"] = ((AliasedValue)oppProductNRR.Entities[0].Attributes["NRR"]).Value; 
                        }

                        step = 10;
                        if (oppProductSetupNet.Entities.Count > 0)
                        {
                            step = 92;
                            opportunityEntity["fdx_totalonetimefees"] = ((AliasedValue)oppProductSetupNet.Entities[0].Attributes["SetupNet"]).Value;
                        }
                        step = 11;
                        service.Update(opportunityEntity);
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the OpportunityProduct_Create plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("OpportunityProduct_Create: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }
    }
}
