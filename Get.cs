using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Organization;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Xml.Linq;

namespace PaginatedReportBuilder
{
    public partial class Get
    {
        public static List<EntityMetadata> GetEntities(IOrganizationService Service)
        {
            List<EntityMetadata> entities = null;
            try
            {
                var request = new RetrieveAllEntitiesRequest
                {
                    EntityFilters = EntityFilters.All
                };

                var response = (RetrieveAllEntitiesResponse)Service.Execute(request);
                entities = response.EntityMetadata.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return entities;
        }

        public static List<Entity> GetEntityForms(int entityTypeCode, IOrganizationService Service)
        {
            var query = new QueryExpression("systemform");
            query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, entityTypeCode);
            query.ColumnSet.AddColumn("name");
            query.ColumnSet.AddColumn("formxml");

            return Service.RetrieveMultiple(query).Entities.ToList();
        }

        public static List<string> GetFormFields(XDocument formDoc, string entitySelected, IOrganizationService Service)
        {
            var controlList = formDoc.Descendants("control").Where(x => x.Attribute("indicationOfSubgrid") == null && x.Descendants("QuickForms").ToList().Count == 0).ToList();
            var fieldList = new List<string>();

            RetrieveEntityRequest req = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entitySelected
            };
            EntityMetadata entitymeta = (Service.Execute(req) as RetrieveEntityResponse).EntityMetadata;

            if (entitymeta != null)
            {
                string field;
                foreach (var control in controlList)
                {
                    field = control.Attribute("id").Value.Replace("header_", "");
                    if (!fieldList.Contains(field) &&
                        entitymeta.Attributes.Where(x => x.LogicalName == field).ToList().Count > 0) // check if real field (could be a web resource)
                    {
                        fieldList.Add(field);
                    }
                }
            }

            return fieldList;
        }

        public static Uri GetOrganizationUrl(IOrganizationService Service)
        {
            var request = new RetrieveCurrentOrganizationRequest();
            var organizationResponse = (RetrieveCurrentOrganizationResponse)Service.Execute(request);

            var uriString = organizationResponse.Detail.Endpoints[EndpointType.WebApplication];
            return new Uri(uriString);
        }

        public static ViewMeta GetViewFields(string viewId, IOrganizationService Service)
        {
            ViewMeta viewMeta = null;

            QueryExpression viewQuery = new QueryExpression
            {
                ColumnSet = new ColumnSet("savedqueryid", "fetchxml", "returnedtypecode"),
                EntityName = "savedquery",
                Criteria = new FilterExpression
                {
                    Conditions ={
                        new ConditionExpression
                        {
                            AttributeName = "savedqueryid",
                            Operator = ConditionOperator.Equal,
                            Values = { viewId }
                        }
                    }
                }
            };

            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = viewQuery };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)Service.Execute(retrieveSavedQueriesRequest);
            Entity retrievedQuery = retrieveSavedQueriesResponse.EntityCollection.Entities.First();

            if (retrievedQuery != null)
            {
                RetrieveEntityRequest req = new RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = retrievedQuery["returnedtypecode"].ToString()
                };
                EntityMetadata entitymeta = (Service.Execute(req) as RetrieveEntityResponse).EntityMetadata;

                if (entitymeta != null) {
                    
                    XDocument viewDoc = XDocument.Parse(retrievedQuery["fetchxml"].ToString());
                    List<XElement> viewColumns = viewDoc.Descendants("attribute").ToList();
                    List<string> fieldList = new List<string>();
                    string field;
                    foreach (XElement column in viewColumns)
                    {
                        field = column.Attribute("name").Value;

                        if (!fieldList.Contains(field) &&
                            entitymeta.Attributes.Where(x => x.LogicalName == field).ToList().Count > 0) // check if real field (could be from related record)
                        {
                            fieldList.Add(field);
                        }
                    }
                    viewMeta = new ViewMeta(fieldList, viewDoc);
                }
            }

            return viewMeta;
        }
    }
}
