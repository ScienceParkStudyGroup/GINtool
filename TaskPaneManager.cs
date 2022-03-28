using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;

namespace GINtool
{
    // from https://stackoverflow.com/questions/19160158/customtaskpane-in-excel-doesnt-appear-in-new-workbooks
    class TaskPaneManager
    {
        static Dictionary<string, CustomTaskPane> _createdPanes = new Dictionary<string, CustomTaskPane>();


        /// <summary>
        /// Gets the taskpane by name (if exists for current excel window then returns existing instance, otherwise uses taskPaneCreatorFunc to create one). 
        /// </summary>
        /// <param name="taskPaneId">Some string to identify the taskpane</param>
        /// <param name="taskPaneTitle">Display title of the taskpane</param>
        /// <param name="taskPaneCreatorFunc">The function that will construct the taskpane if one does not already exist in the current Excel window.</param>
        public static CustomTaskPane GetTaskPane(string taskPaneId, string taskPaneTitle, Func<GINtaskpane> taskPaneCreatorFunc, GINtaskpane.UpdateButtonStatus updateButtonStatus)
        {
            string key = string.Format("{0}({1})", taskPaneId, Globals.ThisAddIn.Application.Hwnd);
            if (!_createdPanes.ContainsKey(key))
            {
                var pane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPaneCreatorFunc(), taskPaneTitle);
                ((GINtaskpane)pane.Control).updateButtonStatus = updateButtonStatus;
                pane.VisibleChanged += new System.EventHandler(TaskPane_VisibleChangedEvent);
                pane.Width = 300;
                _createdPanes[key] = pane;
            }
            return _createdPanes[key];
        }

        // trigger event if window is closed manually
        private static void TaskPane_VisibleChangedEvent(object sender, EventArgs e)
        {
            if (sender != null)
            {
                CustomTaskPane aPane = (CustomTaskPane)sender;
                // sync status with button
                if (aPane != null & aPane.Control is GINtool.GINtaskpane)
                    ((GINtaskpane)aPane.Control).updateButtonStatus(aPane.Visible);
            }

        }
    }

}
