using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using XrmToolBox.Extensibility;

namespace PaginatedReportBuilder
{
    public partial class PaginatedReportBuilderControl : PluginControlBase
    {
        private Settings mySettings;
        private List<Entity> forms;
        private List<EntityMetadata> entities;
        private string dataSource;
        private List<DatasetMeta> datasets;
        private readonly double bodyHeight = 9;
        private readonly double bodyWidth = 6.5;
        private readonly double titleHeight = .5;
        private readonly double cellHeight = .225;
        private readonly double cellPadding = .014;
        private int textBoxIndex, tableIndex;
        private string entitySelected;
        private XDocument formDoc, generatedReport;
        
        public PaginatedReportBuilderControl()
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

        private void PaginatedReportGeneratorControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        private void btn_loadEntities_Click(object sender, EventArgs e)
        {
            box_entitySelect.Items.Clear();
            lst_forms.Items.Clear();

            LoadEntities();
        }

        private void box_entitySelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            lst_forms.Items.Clear();

            if (box_entitySelect.SelectedIndex != -1)
            {
                LoadForms(box_entitySelect.SelectedItem.ToString());
            }
        }

        private void lst_forms_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_formxml.Text = "";
            txt_reportxml.Text = "";
            btn_download.Enabled = false;

            string formXml = forms.Find(x => x["name"].ToString() == lst_forms.SelectedItem.ToString())["formxml"].ToString();
            var formDoc = XDocument.Parse(formXml);
            txt_formxml.Text = formDoc.ToString();
            btn_generate.Enabled = true;
        }

        private void btn_generate_Click(object sender, EventArgs e)
        {
            if (lst_forms.SelectedItem != null)
            {
                DisableInputs();
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
                        EnableInputs();
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

        private void LoadEntities()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving entities...",
                Work = (w, e2) =>
                {
                    // This code is executed in another thread
                    entities = Get.GetEntities(Service);

                    w.ReportProgress(-1, "Entities loaded.");
                    e2.Result = 1;
                },
                ProgressChanged = e2 =>
                {
                    SetWorkingMessage(e2.UserState.ToString());
                },
                PostWorkCallBack = e2 =>
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

        private void LoadForms(string entity)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving forms...",
                Work = (w, e) =>
                {
                    // This code is executed in another thread
                    forms = Get.GetEntityForms((int)entities.Find(x => x.LogicalName == entity).ObjectTypeCode, Service);

                    w.ReportProgress(-1, "Forms loaded.");
                    e.Result = 1;
                },
                ProgressChanged = e =>
                {
                    SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = e =>
                {
                    // This code is executed in the main thread
                    foreach (var form in forms)
                    {
                        lst_forms.Items.Add(form["name"]);
                    }

                    EnableInputs();
                },
                AsyncArgument = null,
                // Progress information panel size
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private void DisableInputs()
        {
            box_entitySelect.Enabled = false;
            lst_forms.Enabled = false;
            btn_generate.Enabled = false;
            btn_download.Enabled = false;
        }

        private void EnableInputs()
        {
            box_entitySelect.Enabled = true;
            lst_forms.Enabled = true;
            btn_generate.Enabled = true;
            btn_download.Enabled = true;
        }

        private void GenerateReport(string formXml)
        {
            try
            {
                string connectString = Get.GetOrganizationUrl(Service).ToString();
                int startIdx = connectString.IndexOf("https://") + ("https://").Length;
                int endIdx = connectString.IndexOf(".crm");
                dataSource = connectString.Substring(startIdx, endIdx - startIdx);

                formDoc = XDocument.Parse(formXml);

                List<string> fields = Get.GetFormFields(formDoc, entitySelected, Service);
                List<string> fieldsXml = Build.BuildDatasetFieldsXml(entitySelected, fields, Service);

                datasets = new List<DatasetMeta>();
                // add main prefiltered dataset
                DatasetMeta dataset1 = Build.BuildDataset(entitySelected, dataSource, datasets, entitySelected, fields, fieldsXml, null, null, Service);
                datasets.Add(dataset1);

                List<string> parameters = Build.BuildParameters(entitySelected);
                List<string> pages = GenerateReportItems(formXml);

                // create header
                List<XElement> headerFields = formDoc.Descendants("cell").Where(x => x.Descendants("control").Where(y => y.Attribute("id").Value.Contains("header_")).ToList().Count > 0).ToList();
                List<string> reportHeaderXml = new List<string>();
                string txtLabel, txtValue, reportCell;
                double cellWidth = bodyWidth / headerFields.Count;
                double leftOffset = 0;
                foreach (XElement cell in headerFields)
                {
                    // create text box
                    txtLabel = cell.Element("labels").Element("label").Attribute("description").Value;
                    txtValue = $"=\"&lt;b&gt;{txtLabel}:&lt;/b&gt; \" + First(Fields!{cell.Element("control").Attribute("id").Value.Replace("header_", "")}.Value, \"{entitySelected}\")";

                    reportCell = Build.BuildTextBox(textBoxIndex, txtLabel, txtValue, 0, leftOffset, cellWidth, 10, cellHeight);
                    reportHeaderXml.Add(reportCell);
                    textBoxIndex++;
                    leftOffset += cellWidth;
                }

                string reportXml = $@"<?xml version=""1.0"" encoding=""utf-8""?>
                        <Report MustUnderstand=""df"" xmlns=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition"" xmlns:rd=""http://schemas.microsoft.com/SQLServer/reporting/reportdesigner"" xmlns:df=""http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition/defaultfontfamily"">
                            <df:DefaultFontFamily>Segoe UI</df:DefaultFontFamily>
                            <AutoRefresh>0</AutoRefresh>
                            <DataSources>
                                <DataSource Name=""{dataSource}"">
                                    <ConnectionProperties>
                                        <DataProvider>MSCRMFETCH</DataProvider>
                                        <ConnectString>{connectString}</ConnectString>
                                    </ConnectionProperties>
                                    <rd:SecurityType>DataBase</rd:SecurityType>
                                    <rd:DataSourceID>e4424900-e2a3-4595-8520-4ee22c8be663</rd:DataSourceID>
                                </DataSource>
                            </DataSources>
                            <DataSets>
                                {String.Join("\n", datasets.Select(x => x.fetchxml).ToArray())}
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
                                        <PageHeader>
                                          <Height>{cellHeight}in</Height>
                                          <PrintOnFirstPage>true</PrintOnFirstPage>
                                          <PrintOnLastPage>true</PrintOnLastPage>
                                          <ReportItems>
                                            {String.Join("\n", reportHeaderXml.ToArray())}
                                          </ReportItems>
                                          <Style>
                                            <Border>
                                              <Style>None</Style>
                                            </Border>
                                          </Style>
                                        </PageHeader>
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

        private List<string> GenerateReportItems(string formXml)
        {
            var formDoc = XDocument.Parse(formXml);
            List<string> pages = new List<string>();
            List<string> reportTabElements = new List<string>();
            List<string> viewFields = new List<string>();
            List<string> viewFieldsXml = new List<string>();
            List<XElement> tabList = formDoc.Descendants("tab").ToList();
            List<XElement> columnList, sectionList, rowList, cellList;

            ViewMeta viewMeta;
            DatasetMeta datasetMeta;
            string page, title, reportCell, reportTable, targetEntity, viewId, controlId, txtLabel, txtValue, pageBreak;
            int tabIndex = 1;
            double tabTopOffset = 0.0, topOffset = 0.0, leftOffset, cellWidth, itemHeight = cellHeight;
            bool textBoxAdded = false;
            bool tableAdded = false;
            textBoxIndex = 1;
            tableIndex = 1;

            try
            {                
                foreach (XElement tab in tabList)
                {
                    reportTabElements.Clear();

                    title = tab.Element("labels").Element("label").Attribute("description").Value.Replace("&", " and ").Replace("“", "").Replace("”", "").Replace("\"", "").Replace("'", "");
                    reportCell = Build.BuildTextBox(textBoxIndex++, "", title, 0, 0, bodyWidth, 20, titleHeight);
                    reportTabElements.Add(reportCell);
                    topOffset = titleHeight + (cellPadding * 2);

                    columnList = tab.Descendants("column").ToList();
                    foreach (XElement column in columnList)
                    {
                        cellWidth = bodyWidth;
                        
                        sectionList = column.Descendants("section").ToList();
                        foreach (XElement section in sectionList)
                        {
                            leftOffset = 0;

                            if (section.Attribute("showlabel").Value == "true" &&
                                section.Descendants("row").Descendants("cell")
                                    .Where(x => x.Descendants("QuickForms").ToList().Count == 0)
                                    .Where(x => x.Descendants("WebResourceId").ToList().Count == 0)
                                    .Where(x => x.Descendants().Where(y => y.Name.LocalName.Contains("UClient")).ToList().Count == 0).ToList().Count > 0) // has >= 1 control that is not a quick view, web resource or UClient control (i.e. timeline)
                            {
                                topOffset += cellHeight + (cellPadding * 2); // for padding

                                txtValue = section.Element("labels").Element("label").Attribute("description").Value.Replace("&", " and ").Replace("“", "").Replace("”", "").Replace("\"", "").Replace("'", "");
                                reportCell = Build.BuildTextBox(textBoxIndex++, "", txtValue, topOffset, leftOffset, cellWidth, 15, cellHeight * 1.5);
                                reportTabElements.Add(reportCell);

                                topOffset += (cellHeight * 1.5) + (cellPadding * 2);
                            }
                            rowList = section.Descendants("row").ToList();
                            foreach (XElement row in rowList)
                            {
                                textBoxAdded = false;
                                tableAdded = false;
                                leftOffset = 0;

                                cellList = row.Descendants("cell")
                                            .Where(x => x.Descendants("QuickForms").ToList().Count == 0)
                                            .Where(x => x.Descendants("WebResourceId").ToList().Count == 0)
                                            .Where(x => x.Descendants().Where(y => y.Name.LocalName.Contains("UClient")).ToList().Count == 0).ToList(); // all controls that are not quick views, web resources or UClient control (i.e. timeline)
                                if (cellList.Count > 0)
                                {
                                    cellWidth = (bodyWidth / cellList.Count) - (cellPadding * 2);

                                    foreach (XElement cell in cellList)
                                    {
                                        if (cell.Descendants("control").Where(x => x.Attribute("indicationOfSubgrid") == null).ToList().Count > 0)
                                        {
                                            // create text box
                                            txtLabel = cell.Element("labels").Element("label").Attribute("description").Value.Replace("&", " and ").Replace("“", "").Replace("”", "").Replace("\"", "").Replace("'", "");
                                            txtValue = $"=\"&lt;b&gt;{txtLabel}:&lt;/b&gt; \" + First(Fields!{cell.Element("control").Attribute("id").Value}.Value, \"{entitySelected}\")";

                                            reportCell = Build.BuildTextBox(textBoxIndex++, txtLabel, txtValue, topOffset, leftOffset, cellWidth, 10, cellHeight);
                                            reportTabElements.Add(reportCell);

                                            textBoxAdded = true;
                                            leftOffset += cellWidth + (cellPadding * 2);
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

                                                viewMeta = Get.GetViewFields(viewId, Service);
                                                viewFields = viewMeta.fields;
                                                viewFieldsXml = Build.BuildDatasetFieldsXml(targetEntity, viewFields, Service);

                                                string relationship = cell.Element("control").Element("parameters").Element("RelationshipName").Value;

                                                // pass entire view fetch to preserve filters
                                                datasetMeta = Build.BuildDataset(entitySelected, dataSource, datasets, targetEntity, null, viewFieldsXml, relationship, viewMeta.fetchxml, Service); // BuildDataset(targetEntity, null, viewFieldsXml, relationship, viewMeta.fetchxml);
                                                if (datasetMeta != null)
                                                {
                                                    datasets.Add(datasetMeta);                                                

                                                    // create title box and table
                                                    if (cell.Attribute("showlabel").Value == "true")
                                                    {
                                                        topOffset += cellHeight + (cellPadding * 2); // for padding

                                                        txtValue = cell.Element("labels").Element("label").Attribute("description").Value.Replace("&", " and ").Replace("“", "").Replace("”", "").Replace("\"", "").Replace("'", "");
                                                        reportCell = Build.BuildTextBox(textBoxIndex++, "", txtValue, topOffset, leftOffset, cellWidth, 15, titleHeight);
                                                        reportTabElements.Add(reportCell);

                                                        topOffset += titleHeight + (cellPadding * 2);
                                                    }

                                                    reportTable = BuildTable(targetEntity, viewFields, topOffset, leftOffset, cellWidth, datasetMeta.name);
                                                    reportTabElements.Add(reportTable);
                                                    tableIndex++;

                                                    tableAdded = true;
                                                    leftOffset += cellWidth;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (textBoxAdded) topOffset += cellHeight + (cellPadding * 2);
                                else if (tableAdded) topOffset += (cellHeight + (cellPadding * 2)) * 2; // tables will have 2 rows
                            }
                        }
                    }

                    pageBreak = tabIndex == 1 ? "None" : "Start";
                    page = $@"<Rectangle Name=""Rectangle{tabIndex++}"">
                            <ReportItems>
                                {String.Join("\n", reportTabElements.ToArray())}
                            </ReportItems>
                            <PageBreak>
                                <BreakLocation>{pageBreak}</BreakLocation>
                            </PageBreak>
                            <KeepTogether>true</KeepTogether>
                            <Top>{tabTopOffset}in</Top>
                            <Height>{topOffset}in</Height>
                            <Width>{bodyWidth}in</Width>
                            <ZIndex>1</ZIndex>
                            <Style>
                                <Border>
                                    <Style>None</Style>
                                </Border>
                            </Style>
                        </Rectangle>";

                    pages.Add(page);
                    tabTopOffset += topOffset;
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return pages;
        }

        private string BuildTable(string entity, List<string> viewFields, double top, double left, double width, string datasetName)
        {
            string columns = "", members = "", headerColumns = "", bodyColumns = "";
            double colWidth = width / (viewFields.Count - 1); // subtract 1 to account for unique identifier column

            RetrieveEntityRequest req = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entity
            };
            EntityMetadata entitymeta = (Service.Execute(req) as RetrieveEntityResponse).EntityMetadata;

            string label;
            foreach (string field in viewFields)
            {
                if (field == entity + "id") continue; // unique identifier

                columns += $@"<TablixColumn>
                              <Width>{colWidth}in</Width>
                            </TablixColumn>";
                members += "<TablixMember />";

                label = $"&lt;b&gt;{entitymeta.Attributes.Where(x => x.LogicalName == field).First().DisplayName.UserLocalizedLabel.Label.Replace("&", " and ").Replace("“", "").Replace("”", "").Replace("\"", "").Replace("'", "")}&lt;/b&gt;";
                headerColumns += Build.BuildTableCell(textBoxIndex++, label, "");

                bodyColumns += Build.BuildTableCell(textBoxIndex++, label, field);
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
                      <Group Name=""{datasetName}_details"">
                        <GroupExpressions>
                          <GroupExpression>=Fields!{entity}id.Value</GroupExpression>
                        </GroupExpressions>
                      </Group>
                    </TablixMember>
                  </TablixMembers>
                </TablixRowHierarchy>
                <DataSetName>{datasetName}</DataSetName>
                <Top>{top}in</Top>
                <Left>{left}in</Left>
                <Height>{cellHeight * 2}in</Height>
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
    }
}