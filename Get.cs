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
using System.ServiceModel;
using System.Xml.Linq;

namespace PaginatedReportGenerator
{
    public partial class Get
    {
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

        public static List<EntityMetadata> GetSolutionEntities(string SolutionUniqueName, IOrganizationService Service)
        {
            // get solution components for solution unique name
            QueryExpression componentsQuery = new QueryExpression
            {
                EntityName = "solutioncomponent",
                ColumnSet = new ColumnSet("objectid"),
                Criteria = new FilterExpression(),
            };
            LinkEntity solutionLink = new LinkEntity("solutioncomponent", "solution", "solutionid", "solutionid", JoinOperator.Inner);
            solutionLink.LinkCriteria = new FilterExpression();
            solutionLink.LinkCriteria.AddCondition(new ConditionExpression("uniquename", ConditionOperator.Equal, SolutionUniqueName));
            componentsQuery.LinkEntities.Add(solutionLink);
            componentsQuery.Criteria.AddCondition(new ConditionExpression("componenttype", ConditionOperator.Equal, 1));
            EntityCollection ComponentsResult = Service.RetrieveMultiple(componentsQuery);
            //Get all entities
            RetrieveAllEntitiesRequest AllEntitiesrequest = new RetrieveAllEntitiesRequest()
            {
                EntityFilters = EntityFilters.Entity | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                RetrieveAsIfPublished = true
            };
            RetrieveAllEntitiesResponse AllEntitiesresponse = (RetrieveAllEntitiesResponse)Service.Execute(AllEntitiesrequest);
            //Join entities Id and solution Components Id 
            return AllEntitiesresponse.EntityMetadata.Join(ComponentsResult.Entities.Select(x => x.Attributes["objectid"]), x => x.MetadataId, y => y, (x, y) => x).ToList();
        }

        public static EntityCollection GetSolutions(IOrganizationService Service)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "solution",
                ColumnSet = new ColumnSet(true),
                Criteria = {
                    Conditions = {
                        new ConditionExpression("ismanaged", ConditionOperator.Equal, false)
                    },
                }
            };

            RetrieveMultipleRequest request = new RetrieveMultipleRequest();
            request.Query = query;
            try
            {
                RetrieveMultipleResponse response = (RetrieveMultipleResponse)Service.Execute(request);
                EntityCollection results = response.EntityCollection;
                return results;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new Exception(ex.Message);
            }
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
