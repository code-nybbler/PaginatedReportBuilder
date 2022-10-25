using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace PaginatedReportGenerator
{
    public partial class Build
    {
        public static DatasetMeta BuildDataset(string entitySelected, string dataSource, List<DatasetMeta> datasets, string entity, List<string> fields, List<string> fieldsXml, string relationship, XDocument viewFetchXml, IOrganizationService Service)
        {
            string parameters = "", fetchXml;

            if (viewFetchXml != null) // related entity
            {
                RetrieveEntityRequest req = new RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Relationships,
                    LogicalName = entity
                };
                EntityMetadata entitymeta = (Service.Execute(req) as RetrieveEntityResponse).EntityMetadata;

                if (entitymeta != null)
                {
                    var relToPrimary = entitymeta.ManyToOneRelationships.Where(x => x.SchemaName == relationship).First();

                    if (relToPrimary != null && viewFetchXml.Descendants("link-entity").Where(x => x.Attribute("name").Value.ToString() == entitySelected).ToList().Count == 0)
                    {
                        string linkedEntity = $"\n    <link-entity name=\"{entitySelected}\" alias=\"aa\" link-type=\"inner\" from=\"{entitySelected}id\" to=\"{relToPrimary.ReferencingAttribute}\" enableprefiltering=\"1\">\n\n        <attribute name=\"{entitySelected}id\" />\n\n</link-entity>\n\n";

                        viewFetchXml.Descendants("entity").First().Add(linkedEntity); // add prefiltered link back to main entity
                    }
                }

                fetchXml = viewFetchXml.ToString();
            }
            else // main entity
            {
                // create query expression to retrieve fetchxml
                QueryExpression query = new QueryExpression(entity);
                query.ColumnSet.AddColumns(fields.ToArray());

                QueryExpressionToFetchXmlRequest request = new QueryExpressionToFetchXmlRequest()
                {
                    Query = query
                };
                QueryExpressionToFetchXmlResponse response = (QueryExpressionToFetchXmlResponse)Service.Execute(request);

                parameters = $@"<QueryParameters>
                                    <QueryParameter Name=""CRM_{entity}"">
                                        <Value>=Parameters!CRM_{entity}.Value</Value>
                                    </QueryParameter>
                                </QueryParameters>";

                fetchXml = response.FetchXml.Replace($"entity name=\"{entity}\"", $"entity name=\"{entity}\" enableprefiltering=\"1\"");
            }

            string name = entity;
            int nameIdx = 2;
            while (datasets.Where(x => x.name == name).ToList().Count > 0)
            {
                name = entity + nameIdx++;
            }

            string datasetFetchXml = $@"<DataSet Name=""{entity}"">
                                    <Query>
                                        <DataSourceName>{dataSource}</DataSourceName>
                                        {parameters}
                                        <CommandText>{fetchXml.Replace("<", "&lt;").Replace(">", "&gt;\n")}</CommandText>
                                    </Query>
                                    <Fields>
                                        {String.Join("\n", fieldsXml)}
                                    </Fields>
                                </DataSet>";

            return new DatasetMeta(name, datasetFetchXml);
        }

        public static List<string> BuildDatasetFieldsXml(string entity, List<string> fieldList, IOrganizationService Service)
        {
            List<string> fieldXmlList = new List<string>();

            try
            {
                RetrieveEntityRequest req = new RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = entity
                };
                EntityMetadata entitymeta = (Service.Execute(req) as RetrieveEntityResponse).EntityMetadata;

                if (entitymeta != null)
                {
                    AttributeMetadata fieldmeta;
                    string datatype, convertedType;
                    string[] specialFieldTypes = { "Lookup", "DateTime", "Picklist", "Decimal", "Integer" };
                    foreach (var field in fieldList)
                    {
                        fieldmeta = entitymeta.Attributes.FirstOrDefault(a => a.LogicalName == field);

                        if (fieldmeta != null)
                        {
                            datatype = fieldmeta.AttributeType.ToString();

                            switch (datatype)
                            {
                                case "Integer": convertedType = "Int32"; break;
                                case "Lookup": convertedType = "Guid"; break;
                                case "Memo":
                                case "Picklist": convertedType = "String"; break;
                                default: convertedType = fieldmeta.AttributeType.ToString(); break;
                            }

                            // every datatype becomes a string
                            // 'special' types get additional columns added for value and entity type
                            string fieldXml = $@"<Field Name=""{field}"">
                                                    <DataField>{field}</DataField>
                                                    <rd:TypeName>System.String</rd:TypeName>
                                                </Field>";

                            if (datatype == "Lookup")
                            {
                                fieldXml += $@"<Field Name=""{field}EntityName"">
                                                  <DataField>{field}EntityName</DataField>
                                                  <rd:TypeName>System.String</rd:TypeName>
                                                </Field>";
                            }
                            if (specialFieldTypes.Contains(datatype))
                            {
                                fieldXml += $@"<Field Name=""{field}Value"">
                                                <DataField>{field}Value</DataField>
                                                <rd:TypeName>System.{convertedType}</rd:TypeName>
                                            </Field>";
                            }

                            fieldXmlList.Add(fieldXml);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return fieldXmlList;
        }

        public static List<string> BuildParameters(string entity)
        {
            List<string> parameters = new List<string>();

            try
            {
                string dataset1 = $@"<ReportParameter Name=""CRM_{entity}"">
                                      <DataType>String</DataType>
                                      <DefaultValue>
                                        <Values>
                                          <Value>&lt;fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false""&gt;&lt;entity name=""{entity}""&gt;&lt;all-attributes/&gt;&lt;/entity&gt;&lt;/fetch&gt;</Value>
                                        </Values>
                                      </DefaultValue>
                                      <Prompt>CRM {entity.Replace("_", " ")}</Prompt>
                                    </ReportParameter>";

                parameters.Add(dataset1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return parameters;
        }

        public static string BuildTableCell(int textBoxIndex, string label, string value)
        {
            string val = value == "" ? label : $"=Fields!{value}.Value";

            string cell = $@"<TablixCell>
                                <CellContents>
                                    <Textbox Name=""Textbox{textBoxIndex}"">
                                        <CanGrow>true</CanGrow>
                                        <KeepTogether>true</KeepTogether>
                                        <Paragraphs>
                                        <Paragraph>
                                            <TextRuns>
                                                <TextRun>
                                                    <Value>{val}</Value>
                                                    <MarkupType>HTML</MarkupType>
                                                    <Style />
                                                </TextRun>
                                            </TextRuns>
                                            <Style />
                                        </Paragraph>
                                        </Paragraphs>
                                        <rd:DefaultName>Textbox{textBoxIndex}</rd:DefaultName>
                                        <Style>
                                        <Border>
                                            <Color>LightGrey</Color>
                                            <Style>Solid</Style>
                                        </Border>
                                        <PaddingLeft>.014in</PaddingLeft>
                                        <PaddingRight>.014in</PaddingRight>
                                        <PaddingTop>.014in</PaddingTop>
                                        <PaddingBottom>.014in</PaddingBottom>
                                        </Style>
                                    </Textbox>
                                </CellContents>
                            </TablixCell>";

            return cell;
        }

        public static string BuildTextBox(int textBoxIndex, string label, string value, double top, double left, double width, int fontSize, double height)
        {
            string textBox = $@"<Textbox Name=""Textbox{textBoxIndex}"">
                                    <CanGrow>true</CanGrow>
                                    <KeepTogether>true</KeepTogether>
                                    <Paragraphs>
                                        <Paragraph>
                                        <TextRuns>
                                            <TextRun>
                                                <Label>{label}</Label>
                                                <Value>{value}</Value>
                                                <MarkupType>HTML</MarkupType>
                                                <Style>
                                                  <FontSize>{fontSize}pt</FontSize>
                                                </Style>
                                            </TextRun>
                                        </TextRuns>
                                        <Style />
                                        </Paragraph>
                                    </Paragraphs>
                                    <rd:DefaultName>Textbox{textBoxIndex}</rd:DefaultName>
                                    <Top>{top}in</Top>
                                    <Left>{left}in</Left>
                                    <Height>{height}in</Height>
                                    <Width>{width}in</Width>
                                    <ZIndex>1</ZIndex>
                                    <Style>
                                        <Border>
                                            <Style>None</Style>
                                        </Border>
                                        <PaddingLeft>.014in</PaddingLeft>
                                        <PaddingRight>.014in</PaddingRight>
                                        <PaddingTop>.014in</PaddingTop>
                                        <PaddingBottom>.014in</PaddingBottom>
                                    </Style>
                                </Textbox>";

            return textBox;
        }
    }
}
