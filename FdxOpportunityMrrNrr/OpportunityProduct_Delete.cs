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
    public class OpportunityProduct_Delete : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            int step = 0;

            if (context.InputParameters.Contains("Target"))
            {
                step = 1;
                //Entity OppProductEntity = (Entity)context.InputParameters["Target"];

                step = 2;
                Entity OppProductPreImageEntity = ((context.PreEntityImages != null) && context.PreEntityImages.Contains("oppproductpre")) ? context.PreEntityImages["oppproductpre"] : null;

                step = 3;
                if (OppProductPreImageEntity.LogicalName != "opportunityproduct")
                    return;


                step = 4;
                Entity oppProductAllEntity = new Entity();
                Entity opportunityEntity = new Entity();
                EntityCollection oppProductMRR = new EntityCollection();
                EntityCollection oppProductNRR = new EntityCollection();
                EntityCollection oppProductSetupFee = new EntityCollection();
                Guid opportunityId = Guid.Empty;

                try
                {
                    step = 5;
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    step = 6;
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    //Retrieve Opportunity Id from Opportunity Product....
                    step = 7;
                    if (OppProductPreImageEntity.Attributes.Contains("opportunityid"))
                        opportunityId = ((EntityReference)OppProductPreImageEntity.Attributes["opportunityid"]).Id;

                    step = 8;
                    if (opportunityId != Guid.Empty)
                    {
                        //Get Sum MRR....
                        step = 9;
                        oppProductMRR = CRMQueryExpression.getMRR(opportunityId, service);

                        //Get Sum NRR....
                        step = 10;
                        oppProductNRR = CRMQueryExpression.getNRR(opportunityId, service);

                        //Update opportunity....
                        step = 11;
                        opportunityEntity = new Entity
                        {
                            Id = opportunityId,
                            LogicalName = "opportunity"
                        };

                        step = 12;
                        if (oppProductMRR.Entities.Count > 0)
                            opportunityEntity["fdx_totalmrr"] = ((AliasedValue)oppProductMRR.Entities[0].Attributes["MRR"]).Value;  

                        step = 13;
                        if (oppProductNRR.Entities.Count > 0)
                            opportunityEntity["fdx_totalnrr"] = ((AliasedValue)oppProductNRR.Entities[0].Attributes["NRR"]).Value;

                        step = 15;
                        if (oppProductSetupFee.Entities.Count > 0)
                        {
                            step = 93;
                            opportunityEntity["fdx_totalonetimefees"] = ((AliasedValue)oppProductSetupFee.Entities[0].Attributes["SetupFee"]).Value;
                        }
                        step = 14;
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
