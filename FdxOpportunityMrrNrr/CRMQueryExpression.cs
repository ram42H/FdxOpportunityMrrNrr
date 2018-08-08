using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxOpportunityMrrNrr
{
    class CRMQueryExpression
    {
        string filterAttribute;
        ConditionOperator filterOperator;
        object filterValue;

        public CRMQueryExpression(string _attr, ConditionOperator _opr, object _val)
        {
            this.filterAttribute = _attr;
            this.filterOperator = _opr;
            this.filterValue = _val;
        }
        public static QueryExpression getQueryExpression(string _entityName, ColumnSet _columnSet, CRMQueryExpression[] _exp, LogicalOperator _filterOperator = LogicalOperator.And)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = _entityName;
            query.ColumnSet = _columnSet;
            query.Distinct = false;
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = _filterOperator;
            for (int i = 0; i < _exp.Length; i++)
            {
                query.Criteria.AddCondition(_exp[i].filterAttribute, _exp[i].filterOperator, _exp[i].filterValue);
            }

            return query;
        }

        public static int GetOptionsetValue(IOrganizationService _service, string _optionsetName, string _optionsetSelectedText)
        {
            int optionsetValue = 0;
            try
            {
                RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new RetrieveOptionSetRequest
                    {
                        Name = _optionsetName
                    };

                // Execute the request.
                RetrieveOptionSetResponse retrieveOptionSetResponse =
                    (RetrieveOptionSetResponse)_service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    //If the value matches/....
                    if (optionMetadata.Label.UserLocalizedLabel.Label.ToString() == _optionsetSelectedText)
                    {
                        optionsetValue = (int)optionMetadata.Value;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return optionsetValue;
        }
        public static string GetOptionsSetTextForValue(IOrganizationService service, string entityName, string attributeName, int selectedValue)
        {

            RetrieveAttributeRequest retrieveAttributeRequest = new
            RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };
            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            // Access the retrieved attribute.
            Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)
            retrieveAttributeResponse.AttributeMetadata;// Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
            string selectedOptionLabel = null;
            foreach (OptionMetadata oMD in optionList)
            {
                if (oMD.Value == selectedValue)
                {
                    selectedOptionLabel = oMD.Label.LocalizedLabels[0].Label.ToString();
                    break;
                }
            }
            return selectedOptionLabel;
        }

        public static EntityCollection getMRR(Guid _opportunityId, IOrganizationService _service)
        {
            FetchExpression query;
            EntityCollection oppProduct = new EntityCollection();
            string OppProductQuery = "<fetch top='1' aggregate='true' ><entity name='opportunityproduct' ><attribute name='fdx_mrr' alias='MRR' aggregate='sum' /><filter type='and' ><condition attribute='opportunityid' operator='eq' value='{0}' /><condition attribute='fdx_revenuetype' operator='eq' value='756480000' /></filter></entity></fetch>";

            query = new FetchExpression(string.Format(OppProductQuery, _opportunityId));
            oppProduct = _service.RetrieveMultiple(query);

            return oppProduct;
        }

        public static EntityCollection getNRR(Guid _opportunityId, IOrganizationService _service)
        {
            FetchExpression query;
            EntityCollection oppProduct = new EntityCollection();
            string OppProductQuery = "<fetch top='1' aggregate='true' ><entity name='opportunityproduct' ><attribute name='fdx_nrr' alias='NRR' aggregate='sum' /><filter type='and' ><condition attribute='opportunityid' operator='eq' value='{0}' /><condition attribute='fdx_revenuetype' operator='eq' value='756480001' /></filter></entity></fetch>";

            query = new FetchExpression(string.Format(OppProductQuery, _opportunityId));
            oppProduct = _service.RetrieveMultiple(query);

            return oppProduct;
        }

        public static EntityCollection getSetUpFee(Guid _opportunityId, IOrganizationService _service)
        {
            FetchExpression query;
            EntityCollection oppProd = new EntityCollection();

            string OppProductQuery = "<fetch top='1' aggregate='true' ><entity name='opportunityproduct' ><attribute name='fdx_setupfee' alias='SetupFee' aggregate='sum' /><filter type='and' ><condition attribute='opportunityid' operator='eq' value='{0}' /></filter></entity></fetch>";

            query = new FetchExpression(string.Format(OppProductQuery, _opportunityId));
            oppProd = _service.RetrieveMultiple(query);

            return oppProd;
        }
    }
}
