﻿using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Organization;
using Microsoft.Xrm.Sdk.Query;
using NuGet.Versioning;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;

namespace PaginatedReportGenerator
{
    public partial class PaginatedReportGeneratorControl : PluginControlBase
    {
        private Settings mySettings;
        private List<EntityMetadata> entities;
        private List<Entity> forms;
        List<string> datasets;
        private double bodyHeight = 9, bodyWidth = 6.5, titleHeight = .5, cellHeight = .225;
        private int textBoxIndex, tableIndex;
        private string entitySelected;
        public XDocument formDoc, generatedReport;
        
        public PaginatedReportGeneratorControl()
        {
            InitializeComponent();
        }

        private void PaginatedReportGeneratorControl_Load(object sender, EventArgs e)
        {
            // Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        /// <summary>
        /// This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PaginatedReportGeneratorControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        /// <summary>
        /// This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
        }

        private Uri GetOrganizationUrl()
        {
            var request = new RetrieveCurrentOrganizationRequest();
            var organizationResponse = (RetrieveCurrentOrganizationResponse)Service.Execute(request);

            var uriString = organizationResponse.Detail.Endpoints[EndpointType.WebApplication];
            return new Uri(uriString);
        }

        private void btn_loadentities_Click(object sender, EventArgs e)
        {
            LoadEntities();
        }

        private void disableInputs()
        {
            box_entitySelect.Enabled = false;
            lst_forms.Enabled = false;
            btn_loadentities.Enabled = false;
            btn_generate.Enabled = false;
        }

        private void enableInputs()
        {
            box_entitySelect.Enabled = true;
            lst_forms.Enabled = true;
            btn_loadentities.Enabled = true;
            btn_generate.Enabled = true;
        }

        private void LoadEntities()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving entities...",
                Work = (w, e) =>
                {
                    // This code is executed in another thread
                    entities = GetAllEntityMetadata();                    

                    w.ReportProgress(-1, "Entities loaded.");
                    e.Result = 1;
                },
                ProgressChanged = e =>
                {
                    SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = e =>
                {
                    // This code is executed in the main thread
                    foreach (var entity in entities)
                    {
                        box_entitySelect.Items.Add(entity.LogicalName);
                    }

                    box_entitySelect.Enabled = true;
                },
                AsyncArgument = null,
                // Progress information panel size
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private List<EntityMetadata> GetAllEntityMetadata()
        {
            var request = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.All
            };

            var response = (RetrieveAllEntitiesResponse)Service.Execute(request);
            return response.EntityMetadata.ToList();
        }        

        private void box_entitySelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateFormsList();
        }

        private void PopulateFormsList()
        {
            if (box_entitySelect.SelectedIndex != -1)
            {                
                // This code is executed in another thread
                forms = GetEntityForms((int)entities.Find(x => x.LogicalName == box_entitySelect.SelectedItem.ToString()).ObjectTypeCode);

                foreach(var form in forms)
                {
                    lst_forms.Items.Add(form["name"]);
                }
                enableInputs();
            }
        }

        private List<Entity> GetEntityForms(int entityTypeCode)
        {
            var query = new QueryExpression("systemform");
            query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, entityTypeCode);
            query.ColumnSet.AddColumn("name");
            query.ColumnSet.AddColumn("formxml");

            return Service.RetrieveMultiple(query).Entities.ToList(); ;
        }

        private void lst_forms_SelectedIndexChanged(object sender, EventArgs e)
        {
            string formXml = forms.Find(x => x["name"].ToString() == lst_forms.SelectedItem.ToString())["formxml"].ToString();
            var formDoc = XDocument.Parse(formXml);
            txt_formxml.Text = formDoc.ToString();
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            if (lst_forms.SelectedItem != null)
            {
                disableInputs();
                string formXml = txt_formxml.Text;
                entitySelected = box_entitySelect.SelectedItem.ToString();
                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Generating report...",
                    Work = (w, e2) =>
                    {
                        // This code is executed in another thread
                        GenerateReport(formXml);

                        w.ReportProgress(-1, "Report generated.");
                        e2.Result = 1;
                    },
                    ProgressChanged = e2 =>
                    {
                        SetWorkingMessage(e2.UserState.ToString());
                    },
                    PostWorkCallBack = e2 =>
                    {
                        // This code is executed in the main thread
                        txt_reportxml.Text = generatedReport.ToString();
                        enableInputs();
                    },
                    AsyncArgument = null,
                    // Progress information panel size
                    MessageWidth = 340,
                    MessageHeight = 150
                });
            }
        }

        private void btn_download_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    generatedReport.Save(fbd.SelectedPath + "\\" + entitySelected + ".rdl");

                    MessageBox.Show($"Report saved to {fbd.SelectedPath}\\{entitySelected}.rdl");
                }
            }
            
        }

        private void GenerateReport(string formXml)
        {
            try
            {
                formDoc = XDocument.Parse(formXml);                

                List<string> fields = GetFormFields();
                List<string> fieldsXml = BuildDatasetFieldsXml(entitySelected, fields);

                datasets = new List<string>();
                // add main prefiltered dataset
                string dataset1 = BuildDataset(entitySelected, fields, fieldsXml, null, null);
                datasets.Add(dataset1);

                List<string> parameters = BuildParameters();
                List<string> pages = BuildReportItems(formXml);

                string reportXml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                        <Report MustUnderstand=""df"" xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition"" xmlns:rd=""http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"" xmlns:df=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily"">
                            <df:DefaultFontFamily>Segoe UI</df:DefaultFontFamily>
                            <AutoRefresh>0</AutoRefresh>
                            <DataSources>
                                <DataSource Name=""DataSource1"">
                                    <ConnectionProperties>
                                        <DataProvider>MSCRMFETCH</DataProvider>
                                        <ConnectString>{GetOrganizationUrl()}</ConnectString>
                                    </ConnectionProperties>
                                    <rd:SecurityType>DataBase</rd:SecurityType>
                                    <rd:DataSourceID>e4424900-e2a3-4595-8520-4ee22c8be663</rd:DataSourceID>
                                </DataSource>
                            </DataSources>
                            <DataSets>
                                {String.Join("\n", datasets.ToArray())}
                            </DataSets>
                            <ReportSections>
                                <ReportSection>
                                    <Body>
                                        <ReportItems>
                                            {String.Join("\n", pages.ToArray())}
                                        </ReportItems>
                                        <Height>{bodyHeight}in</Height>
                                        <Style />
                                    </Body>
                                    <Width>{bodyWidth}in</Width>
                                        <Page>
                                            <LeftMargin>1in</LeftMargin>
                                            <RightMargin>1in</RightMargin>
                                            <TopMargin>1in</TopMargin>
                                            <BottomMargin>1in</BottomMargin>
                                        <Style />
                                    </Page>
                                </ReportSection>
                            </ReportSections>
                            <ReportParameters>
                                {String.Join("\n", parameters)}
                            </ReportParameters>
                            <ReportParametersLayout>
                                <GridLayoutDefinition>
                                    <NumberOfColumns>4</NumberOfColumns>
                                    <NumberOfRows>2</NumberOfRows>
                                    <CellDefinitions>
                                        <CellDefinition>
                                            <ColumnIndex>0</ColumnIndex>
                                            <RowIndex>0</RowIndex>
                                            <ParameterName>CRM_{entitySelected}</ParameterName>
                                        </CellDefinition>
                                    </CellDefinitions>
                                </GridLayoutDefinition>
                            </ReportParametersLayout>
                            <rd:ReportUnitType>Inch</rd:ReportUnitType>
                            <rd:ReportID>9b07c0c5-bc44-4226-a0ff-1f1e718f7c22</rd:ReportID>
                        </Report>";

                generatedReport = XDocument.Parse(reportXml);                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private List<string> BuildParameters()
        {
            List<string> parameters = new List<string>();

            try
            {
                string dataset1 = $@"<ReportParameter Name=""CRM_{entitySelected}"">
                                      <DataType>String</DataType>
                                      <DefaultValue>
                                        <Values>
                                          <Value>&lt;fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false""&gt;&lt;entity name=""{entitySelected}""&gt;&lt;all-attributes/&gt;&lt;/entity&gt;&lt;/fetch&gt;</Value>
                                        </Values>
                                      </DefaultValue>
                                      <Prompt>CRM {entitySelected.Replace("_", " ")}</Prompt>
                                    </ReportParameter>";

                parameters.Add(dataset1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return parameters;
        }

        private ViewMeta GetViewFields(string viewId)
        {
            List<string> fieldNames = new List<string>();

            QueryExpression viewQuery = new QueryExpression
            {
                ColumnSet = new ColumnSet("savedqueryid", "fetchxml"),
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
            DataCollection<Entity> retrievedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;

            XDocument viewDoc = null;
            List<XElement> viewFields;
            foreach (Entity ent in retrievedQueries)
            {
                viewDoc = XDocument.Parse(ent.Attributes["fetchxml"].ToString());
                viewFields = viewDoc.Descendants("attribute").ToList();
                foreach (XElement field in viewFields)
                {
                    fieldNames.Add(field.Attribute("name").Value);
                }
            }

            return new ViewMeta(fieldNames, viewDoc);
        }

        private string BuildDataset(string entity, List<string> fields, List<string> fieldsXml, string relationship, XDocument viewFetchXml)
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
                        string linkedEntity = $"\n    <link-entity name=\"{entitySelected}\" alias=\"aa\" link-type=\"inner\" from=\"{relToPrimary.ReferencingAttribute}\" to=\"{entitySelected}id\" enableprefiltering=\"1\">\n\n        <attribute name=\"{entitySelected}id\" />\n\n</link-entity>\n\n";

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

            string dataset = $@"<DataSet Name=""{entity}"">
                                    <Query>
                                        <DataSourceName>DataSource1</DataSourceName>
                                        {parameters}
                                        <CommandText>{fetchXml.Replace("<", "&lt;").Replace(">", "&gt;\n")}</CommandText>
                                    </Query>
                                    <Fields>
                                        {String.Join("\n", fieldsXml)}
                                    </Fields>
                                </DataSet>";

            return dataset;
        }

        private List<string> GetFormFields()
        {
            var controlList = formDoc.Descendants("control").Where(x => x.Attribute("indicationOfSubgrid") == null).ToList();
            var fieldList = new List<string>();

            foreach (var control in controlList)
            {
                fieldList.Add(control.Attribute("id").Value.Replace("header_", ""));
            }

            return fieldList;
        }
            
        private List<string> BuildDatasetFieldsXml(string entity, List<string> fieldList)
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
                            if (specialFieldTypes.Contains(datatype)) {
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

        private string BuildTextBox(string label, string value, double top, double left, double width, int fontSize, double height)
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
                                        <PaddingLeft>2pt</PaddingLeft>
                                        <PaddingRight>2pt</PaddingRight>
                                        <PaddingTop>2pt</PaddingTop>
                                        <PaddingBottom>2pt</PaddingBottom>
                                    </Style>
                                </Textbox>";

            return textBox;
        }

        private string BuildTableCell(string label, string value)
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
                                        <PaddingLeft>2pt</PaddingLeft>
                                        <PaddingRight>2pt</PaddingRight>
                                        <PaddingTop>2pt</PaddingTop>
                                        <PaddingBottom>2pt</PaddingBottom>
                                        </Style>
                                    </Textbox>
                                </CellContents>
                            </TablixCell>";

            textBoxIndex++;

            return cell;
        }

        private string BuildTable(string entity, List<string> viewFields, double top, double left, double width)
        {
            string columns = "", members = "", headerColumns = "", bodyColumns = "";
            double colWidth = width / (viewFields.Count-1); // subtract 1 to account for unique identifier column

            RetrieveEntityRequest req = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entity
            };
            EntityMetadata entitymeta = (Service.Execute(req) as RetrieveEntityResponse).EntityMetadata;

            string label;
            foreach (string field in viewFields) {
                if (field == entity + "id") continue; // unique identifier

                columns += $@"<TablixColumn>
                              <Width>{colWidth}in</Width>
                            </TablixColumn>";
                members += "<TablixMember />";

                label = $"&lt;b&gt;{entitymeta.Attributes.Where(x => x.LogicalName == field).First().DisplayName.UserLocalizedLabel.Label}&lt;/b&gt;";
                headerColumns += BuildTableCell(label, "");

                bodyColumns += BuildTableCell(label, field);
            }

            string table = $@"<Tablix Name=""Tablix{tableIndex}"">
                <TablixBody>
                  <TablixColumns>
                    {columns}
                  </TablixColumns>
                  <TablixRows>
                    <TablixRow>
                      <Height>0.25in</Height>
                      <TablixCells>
                        {headerColumns}
                      </TablixCells>
                    </TablixRow>
                    <TablixRow>
                      <Height>0.25in</Height>
                      <TablixCells>
                        {bodyColumns}
                      </TablixCells>
                    </TablixRow>
                  </TablixRows>
                </TablixBody>
                <TablixColumnHierarchy>
                  <TablixMembers>
                    {members}
                  </TablixMembers>
                </TablixColumnHierarchy>
                <TablixRowHierarchy>
                  <TablixMembers>
                    <TablixMember>
                      <KeepWithGroup>After</KeepWithGroup>
                    </TablixMember>
                    <TablixMember>
                      <Group Name=""{entity}_details"">
                        <GroupExpressions>
                          <GroupExpression>=Fields!{entity}id.Value</GroupExpression>
                        </GroupExpressions>
                      </Group>
                    </TablixMember>
                  </TablixMembers>
                </TablixRowHierarchy>
                <DataSetName>{entity}</DataSetName>
                <Top>{top}in</Top>
                <Left>{left}in</Left>
                <Height>0.5in</Height>
                <Width>{width}in</Width>
                <ZIndex>4</ZIndex>
                <Style>
                  <Border>
                    <Style>None</Style>
                  </Border>
                </Style>
              </Tablix>";

            return table;
        }

        private List<string> BuildReportItems(string formXml)
        {
            var formDoc = XDocument.Parse(formXml);
            List<string> pages = new List<string>();
            List<string> reportTabElements = new List<string>();
            List<string> viewFields = new List<string>();
            List<string> viewFieldsXml = new List<string>();
            List<XElement> tabList = formDoc.Descendants("tab").ToList();
            List<XElement> columnList, sectionList, rowList, cellList;

            ViewMeta viewMeta;
            string page, title, reportCell, reportTable, targetEntity, viewId, dataset, controlId, txtLabel, txtValue;
            int tabIndex = 1;
            double tabTopOffset = 0.0, topOffset = 0.0, leftOffset, cellWidth, itemHeight = cellHeight;
            bool itemAdded = false;
            textBoxIndex = 1;
            tableIndex = 1;

            try
            {                
                foreach (XElement tab in tabList)
                {
                    tabTopOffset = 9 * (tabIndex-1);
                    topOffset = titleHeight;
                    reportTabElements.Clear();

                    title = tab.Element("labels").Element("label").Attribute("description").Value;
                    columnList = tab.Descendants("column").ToList();
                    foreach (XElement column in columnList)
                    {
                        cellWidth = Double.Parse(column.Attribute("width").Value.Replace("%", ""))/100.0 * bodyWidth;
                        sectionList = column.Descendants("section").ToList();
                        foreach (XElement section in sectionList)
                        {
                            leftOffset = 0;
                            if (section.Attribute("showlabel").Value == "true")
                            {
                                topOffset += cellHeight;
                                txtValue = section.Element("labels").Element("label").Attribute("description").Value;
                                reportCell = BuildTextBox("", txtValue, topOffset, leftOffset, cellWidth, 15, cellHeight*1.5);
                                reportTabElements.Add(reportCell);
                                textBoxIndex++;
                                topOffset += titleHeight;
                            }
                            rowList = section.Descendants("row").ToList();
                            foreach (XElement row in rowList)
                            {
                                leftOffset = 0;
                                cellList = row.Descendants("cell").ToList();
                                foreach (XElement cell in cellList)
                                {
                                    itemAdded = false;
                                    if (cell.Descendants("control").Where(x => x.Attribute("indicationOfSubgrid") == null).ToList().Count > 0)
                                    {
                                        itemHeight = cellHeight;
                                        // create text box
                                        txtLabel = cell.Element("labels").Element("label").Attribute("description").Value;
                                        txtValue = $"=\"&lt;b&gt;{txtLabel}:&lt;/b&gt; \" + First(Fields!{cell.Element("control").Attribute("id").Value}.Value, \"{entitySelected}\")";

                                        reportCell = BuildTextBox(txtLabel, txtValue, topOffset, leftOffset, cellWidth, 10, cellHeight);
                                        reportTabElements.Add(reportCell);
                                        textBoxIndex++;

                                        itemAdded = true;
                                    }
                                    else if (cell.Descendants("control").ToList().Count > 0)
                                    {
                                        controlId = "";
                                        if (cell.Element("control").Attribute("uniqueid") != null)
                                        {
                                            controlId = cell.Element("control").Attribute("uniqueid").Value;
                                        }

                                        if (controlId == "" || formDoc.Descendants("controlDescription").Where(x => x.Attribute("forControl").Value == controlId).ToList().Count == 0)
                                        {
                                            targetEntity = cell.Element("control").Element("parameters").Element("TargetEntityType").Value;
                                            viewId = cell.Element("control").Element("parameters").Element("ViewId").Value;

                                            viewMeta = GetViewFields(viewId);
                                            viewFields = viewMeta.fields;
                                            viewFieldsXml = BuildDatasetFieldsXml(targetEntity, viewFields);

                                            string relationship = cell.Element("control").Element("parameters").Element("RelationshipName").Value;
                                            
                                            // pass entire view fetch to preserve filters
                                            dataset = BuildDataset(targetEntity, null, viewFieldsXml, relationship, viewMeta.fetchxml);
                                            datasets.Add(dataset);

                                            // create title box and table
                                            if (cell.Attribute("showlabel").Value == "true")
                                            {
                                                topOffset += cellHeight;
                                                txtValue = cell.Element("labels").Element("label").Attribute("description").Value;
                                                reportCell = BuildTextBox("", txtValue, topOffset, leftOffset, cellWidth, 15, titleHeight);
                                                reportTabElements.Add(reportCell);
                                                textBoxIndex++;
                                                topOffset += titleHeight;
                                            }

                                            itemHeight = cellHeight * 2;
                                            reportTable = BuildTable(targetEntity, viewFields, topOffset, leftOffset, cellWidth);
                                            reportTabElements.Add(reportTable);
                                            tableIndex++;

                                            itemAdded = true;
                                        }
                                    }
                                    if (itemAdded == true) leftOffset += cellWidth;
                                }
                                if (itemAdded == true) topOffset += itemHeight;
                            }
                        }
                    }

                    page = $@"<Rectangle Name=""Rectangle{tabIndex++}"">
                            <ReportItems>
                                <Textbox Name=""Textbox{textBoxIndex}"">
                                <CanGrow>true</CanGrow>
                                <KeepTogether>true</KeepTogether>
                                <Paragraphs>
                                    <Paragraph>
                                    <TextRuns>
                                        <TextRun>
                                        <Value>{title}</Value>
                                        <Style>
                                            <FontSize>20pt</FontSize>
                                        </Style>
                                        </TextRun>
                                    </TextRuns>
                                    <Style />
                                    </Paragraph>
                                </Paragraphs>
                                <rd:DefaultName>Textbox{textBoxIndex++}</rd:DefaultName>
                                <Height>{titleHeight}in</Height>
                                <Width>{bodyWidth}in</Width>
                                <Style>
                                    <Border>
                                    <Style>None</Style>
                                    </Border>
                                    <PaddingLeft>2pt</PaddingLeft>
                                    <PaddingRight>2pt</PaddingRight>
                                    <PaddingTop>2pt</PaddingTop>
                                    <PaddingBottom>2pt</PaddingBottom>
                                </Style>
                                </Textbox>
                                {String.Join("\n", reportTabElements.ToArray())}
                            </ReportItems>
                            <PageBreak>
                                <BreakLocation>End</BreakLocation>
                            </PageBreak>
                            <KeepTogether>true</KeepTogether>
                            <Top>{tabTopOffset}in</Top>
                            <Height>{bodyHeight}in</Height>
                            <Width>{bodyWidth}in</Width>
                            <ZIndex>1</ZIndex>
                            <Style>
                                <Border>
                                    <Style>None</Style>
                                </Border>
                            </Style>
                        </Rectangle>";

                    pages.Add(page);
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return pages;
        }
    }
}