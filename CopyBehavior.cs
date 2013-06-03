using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using HPMSdk;

using Hansoft.Jean.Behavior;
using Hansoft.ObjectWrapper;
using Hansoft.ObjectWrapper.CustomColumnValues;

namespace Hansoft.Jean.Behavior.CopyBehavior
{
    public class CopyBehavior : AbstractBehavior
    {
        class ColumnMapping
        {
            internal ColumnDefinition SourceColumn;
            internal ColumnDefinition TargetColumn;
        }

        class ColumnDefinition
        {
            internal bool IsCustomColumn;
            internal string CustomColumnName;
            internal HPMProjectCustomColumnsColumn CustomColumn; 
            internal EHPMProjectDefaultColumn DefaultColumnType; 
        }

        private string title;

        private string targetProjectName;
        private string targetView;
        EHPMReportViewType targetViewType;
        private string targetFind;
        private Project targetProject;
        private ProjectView targetProjectView;

        private string sourceProjectName;
        private string sourceView;
        EHPMReportViewType sourceViewType;
        private string sourceFind;
        private Project sourceProject;
        private ProjectView sourceProjectView;

        private List<ColumnMapping> columnMappings;

        private bool changeImpact;
        
        public CopyBehavior(XmlElement configuration)
            : base(configuration) 
        {
            targetProjectName = GetParameter("TargetProject");
            targetView = GetParameter("TargetView");
            targetViewType = GetViewType(targetView);
            targetFind = GetParameter("TargetFind");
            sourceProjectName = GetParameter("SourceProject");
            sourceView = GetParameter("SourceView");
            sourceViewType = GetViewType(sourceView);
            sourceFind = GetParameter("SourceFind");
            columnMappings = new List<ColumnMapping>();

            foreach (XmlElement cEl in configuration.GetElementsByTagName("ColumnMapping"))
            {
                ColumnMapping columnMapping = new ColumnMapping();
                columnMapping.SourceColumn = GetColumnDefinition(cEl, "Source");
                columnMapping.TargetColumn = GetColumnDefinition(cEl, "Target");
                columnMappings.Add(columnMapping);
            }
            title = "CopyBehavior: " + configuration.InnerText;
        }

        public override string Title
        {
            get { return title; }
        }

        // TODO: Move the ColumnDefinition  class and its parsing to a separate class in the Behavior namespace.
        private ColumnDefinition GetColumnDefinition(XmlElement parent, string tagName)
        {
            ColumnDefinition colDef = new ColumnDefinition();
            XmlElement el = (XmlElement)((XmlElement)parent.GetElementsByTagName(tagName)[0]).ChildNodes[0];
            switch (el.Name)
            {
                case ("CustomColumn"):
                    colDef.IsCustomColumn = true;
                    colDef.CustomColumnName = el.GetAttribute("Name");
                    break;
                case ("Risk"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.Risk;
                    break;
                case ("Priority"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.BacklogPriority;
                    break;
                case ("EstimatedDays"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.EstimatedIdealDays;
                    break;
                case ("Category"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.BacklogCategory;
                    break;
                case ("Points"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.ComplexityPoints;
                    break;
                case ("Status"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.ItemStatus;
                    break;
                case ("Confidence"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.Confidence;
                    break;
                case ("Hyperlink"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.Hyperlink;
                    break;
                case ("Name"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.ItemName;
                    break;
                case ("WorkRemaining"):
                    colDef.IsCustomColumn = false;
                    colDef.DefaultColumnType = EHPMProjectDefaultColumn.WorkRemaining;
                    break;
                default:
                    throw new ArgumentException("Unknown column type specified in Copy behavior : " + el.Name);
            }
            return colDef;
        }

        public override void Initialize()
        {
            targetProject = HPMUtilities.FindProject(targetProjectName);
            if (targetProject == null)
                throw new ArgumentException("Could not find project:" + targetProjectName);
            if (targetViewType == EHPMReportViewType.AgileBacklog)
                targetProjectView = targetProject.ProductBacklog;
            else if (targetViewType == EHPMReportViewType.AllBugsInProject)
                targetProjectView = targetProject.BugTracker;
            else
                targetProjectView = targetProject.Schedule;

            sourceProject = HPMUtilities.FindProject(sourceProjectName);
            if (sourceProject == null)
                throw new ArgumentException("Could not find project:" + sourceProjectName);
            if (sourceViewType == EHPMReportViewType.AgileBacklog)
                sourceProjectView = sourceProject.ProductBacklog;
            else if (targetViewType == EHPMReportViewType.AllBugsInProject)
                sourceProjectView = sourceProject.BugTracker;
            else
                sourceProjectView = sourceProject.Schedule;

            for (int i = 0; i < columnMappings.Count; i += 1 )
            {
                ColumnMapping colDef = columnMappings[i];
                if (colDef.SourceColumn.IsCustomColumn)
                    colDef.SourceColumn.CustomColumn = ResolveCustomColumn(sourceProjectView, colDef.SourceColumn);
                if (colDef.TargetColumn.IsCustomColumn)
                    colDef.TargetColumn.CustomColumn = ResolveCustomColumn(targetProjectView, colDef.TargetColumn);
            }
            DoUpdate();
        }

        // TODO: Move to the ColumnDefinition class (to be...) in the Behavior namespace
        private HPMProjectCustomColumnsColumn ResolveCustomColumn(ProjectView projectView, ColumnDefinition colDef)
        {
            HPMProjectCustomColumnsColumn customColumn = projectView.GetCustomColumn(colDef.CustomColumnName);
            if (customColumn == null)
                throw new ArgumentException("Could not find custom column:" + colDef.CustomColumnName);
            else
                return customColumn;
        }

        // TODO: Move to helper in the Behavior namespace
        private EHPMReportViewType GetViewType(string viewType)
        {
            switch (viewType)
            {
                case ("Agile"):
                    return EHPMReportViewType.AgileMainProject;
                case ("Scheduled"):
                    return EHPMReportViewType.ScheduleMainProject;
                case ("Bugs"):
                    return EHPMReportViewType.AllBugsInProject;
                case ("Backlog"):
                    return EHPMReportViewType.AgileBacklog;
                default:
                    throw new ArgumentException("Unsupported View Type: " + viewType);

            }
        }

        private void DoUpdate()
        {
            List<Task> targetItems = targetProjectView.Find(targetFind);
            List<Task> sourceItems = sourceProjectView.Find(targetFind);
            foreach (Hansoft.ObjectWrapper.Task targetTask in targetItems)
            {
                Hansoft.ObjectWrapper.Task sourceTask = sourceItems.FirstOrDefault(s => s.LinkedTasks.Contains(targetTask));
                if (sourceTask != null)
                {
                    foreach (ColumnMapping mapping in columnMappings)
                    {
                        // TODO: Simplify the sequence below
                        if (mapping.SourceColumn.IsCustomColumn && mapping.TargetColumn.IsCustomColumn && (mapping.SourceColumn.CustomColumn.m_Type == mapping.TargetColumn.CustomColumn.m_Type))
                        {
                            targetTask.SetCustomColumnValue(mapping.TargetColumn.CustomColumn, CustomColumnValue.FromInternalValue(targetTask, mapping.TargetColumn.CustomColumn, sourceTask.GetCustomColumnValue(mapping.SourceColumn.CustomColumn).InternalValue));
                        }
                        else
                        {
                            object sourceValue;
                            if (mapping.SourceColumn.IsCustomColumn)
                                sourceValue = sourceTask.GetCustomColumnValue(mapping.SourceColumn.CustomColumn);
                            else
                                sourceValue = sourceTask.GetDefaultColumnValue(mapping.SourceColumn.DefaultColumnType);
                            if (mapping.TargetColumn.IsCustomColumn)
                            {
                                string endUserString;
                                if (sourceValue is float || sourceValue is double)
                                    endUserString = String.Format(new System.Globalization.CultureInfo("en-US"), "{0:F1}", sourceValue);
                                else
                                    endUserString = sourceValue.ToString();
                                targetTask.SetCustomColumnValue(mapping.TargetColumn.CustomColumn, endUserString);
                            }
                            else
                                targetTask.SetDefaultColumnValue(mapping.TargetColumn.DefaultColumnType, sourceValue);
                        }
                    }
                }
            }
        }

        public override void OnBeginProcessBufferedEvents(EventArgs e)
        {
            changeImpact = false;
        }

        public override void OnTaskChange(TaskChangeEventArgs e)
        {
            // TODO: Optimize by checking which columns has changed
            if (Task.GetTask(e.Data.m_TaskID).MainProjectID.m_ID == sourceProject.Id)
            {
                if (!BufferedEvents)
                    DoUpdate();
                else
                    changeImpact = true;
            }
        }

        public override void OnTaskChangeCustomColumnData(TaskChangeCustomColumnDataEventArgs e)
        {
            // TODO: Optimize by checking which columns has changed
            if (Task.GetTask(e.Data.m_TaskID).MainProjectID.m_ID == sourceProject.Id)
            {
                if (!BufferedEvents)
                    DoUpdate();
                else
                    changeImpact = true;
            }
        }

        public override void OnEndProcessBufferedEvents(EventArgs e)
        {
            if (BufferedEvents && changeImpact)
                DoUpdate();
        }
    }
}
