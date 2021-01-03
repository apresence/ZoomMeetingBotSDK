namespace ZoomMeetngBotSDK
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Windows.Automation;
    using System.Windows.Forms;
    using global::ZoomMeetngBotSDK.Interop.HostApp;
    using global::ZoomMeetngBotSDK.Utils;

    internal class UIATools
    {
        public static string AERuntimeIDToString(int[] rid)
        {
            return string.Join(",", rid);
        }

        /// <summary>
        /// Returns a string representation of Automation Element that is useful for logging purposes.
        /// </summary>
        public static string AEToString(AutomationElement ae)
        {
            if (ae == null)
            {
                return "(null)";
            }

            try
            {
                var l = new List<string>();
                var aei = ae.Current;

                // Based on experimentation, the following AEI's we don't care about:
                //   aei.AcceleratorKey         Never set
                //   aei.FrameworkId            Always "Win32"
                //   aei.HelpText               Never set
                //   aei.IsEnabled              Always true
                //   aei.IsControlElement       Always true
                //   aei.IsOffscreen            Always false
                //   aei.IsPassword             Not relevant
                //   aei.IsRequiredForForm      Not relevant
                //   aei.ItemStatus             Never set
                //   aei.ItemType               Never set
                //   aei.LabeledBy              Never set
                //   aei.LocalizedControlType   Already available as ControlType
                //   aei.Orientation            Always 0

                l.Add(aei.ControlType.ProgrammaticName.Split('.')[1]);
                l.Add(Global.repr(aei.Name));
                if ((aei.AutomationId != null) && (aei.AutomationId.Length > 0))
                {
                    l.Add("AutomationId:" + Global.repr(aei.AutomationId));
                }

                if (!aei.BoundingRectangle.IsEmpty)
                {
                    l.Add("{" + aei.BoundingRectangle.ToString() + "}");
                }
                //if (aei.IsEnabled) l.Add("Enabled");
                if (aei.IsContentElement)
                {
                    l.Add("Content");
                }
                //if (aei.IsControlElement) l.Add("Control");
                if (aei.IsKeyboardFocusable)
                {
                    l.Add("KeyboardFocusable");
                }

                if (aei.HasKeyboardFocus)
                {
                    l.Add("KeyboardFocus");
                }

                if ((aei.AccessKey != null) && (aei.AccessKey.Length > 0))
                {
                    l.Add("AccessKey:" + Global.repr(aei.AccessKey));
                }

                if ((aei.ClassName != null) && (aei.ClassName.Length > 0))
                {
                    l.Add("ClassName:" + Global.repr(aei.ClassName));
                }

                if (aei.NativeWindowHandle != 0)
                {
                    l.Add("hWnd:" + aei.NativeWindowHandle.ToString());
                }

                if (aei.ProcessId != 0)
                {
                    l.Add("pid:" + aei.ProcessId.ToString());
                }

                var p = new List<string>();
                foreach (var pattern in ae.GetSupportedPatterns())
                {
                    p.Add(pattern.ProgrammaticName.Split('.')[0]);
                }

                if (p.Count > 0)
                {
                    l.Add("patterns:[" + string.Join(",", p) + "]");
                }

                l.Add("rids:" + Global.repr(ae.GetRuntimeId()));

                return string.Join(" ", l);
            }
            catch (ElementNotAvailableException)
            {
                return "(ElementNotAvailableException)";
            }
        }

        public static string WalkRawElementsToString(AutomationElement parentElement, bool force = false)
        {
            if (!(Global.cfg.EnableWalkRawElementsToString || force))
            {
                return "(disabled)";
            }

            TreeNode rootNode = new TreeNode();
            WalkRawElements(parentElement, rootNode);
            return TreeNodeToString(rootNode);
        }

        public static void LogSupportedPatterns(AutomationElement ae)
        {
            Global.hostApp.Log(LogType.DBG, "Supported patterns for {0} {1}:", Global.repr(ae.Current.LocalizedControlType), Global.repr(ae.Current.Name));

            List<string> l = new List<string>();

            var aps = ae.GetSupportedPatterns();
            foreach (AutomationPattern ap in aps)
            {
                l.Add(ap.ProgrammaticName);
            }
            Global.hostApp.Log(LogType.DBG, "  {0}", l.Count == 0 ? "None" : string.Join(",", l));
        }

        public static void LogSiblings(AutomationElement ae)
        {
            Global.hostApp.Log(LogType.DBG, "Siblings:");
            var nodeAttendee = TreeWalker.ContentViewWalker.GetNextSibling(ae);
            while (nodeAttendee != null)
            {
                Global.hostApp.Log(LogType.DBG, "  {0}", ZMBUtils.GetObjStrs(nodeAttendee.Current));
                LogSupportedPatterns(nodeAttendee);
                nodeAttendee = TreeWalker.RawViewWalker.GetNextSibling(nodeAttendee);
            }
        }

        public static void LogChildren(AutomationElement ae)
        {
            Global.hostApp.Log(LogType.DBG, "Children:");
            var nodeAttendee = TreeWalker.ContentViewWalker.GetFirstChild(ae);
            while (nodeAttendee != null)
            {
                Global.hostApp.Log(LogType.DBG, "  {0}", ZMBUtils.GetObjStrs(nodeAttendee.Current));
                LogSupportedPatterns(nodeAttendee);
                nodeAttendee = TreeWalker.RawViewWalker.GetNextSibling(nodeAttendee);
            }
        }

        /// <summary>
        /// Returns abbreviated automation event name.  Ex: "AutomationEvent.AutomationFocusChangedEvent" -> "Focus".
        /// </summary>
        public static string GetEventShortName(AutomationEventArgs evt)
        {
            return evt.EventId.ProgrammaticName.Split('.')[1].Replace("Automation", string.Empty).Replace("Changed", string.Empty).Replace("Event", string.Empty);
        }

        /// <summary>
        /// Returns abbreviated control type name.  Ex: "ControlType.ListItem" -> "ListItem".
        /// </summary>
        public static string GetControlTypeShortName(ControlType ct)
        {
            return ct.ProgrammaticName.ToString().Split('.')[1];
        }

        public static string GetText(AutomationElement element)
        {
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                // often there is an extra '\r' hanging off the end.
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r');
            }
            else
            {
                return element.Current.Name;
            }
        }

        private static void WalkRawElements(AutomationElement rootElement, TreeNode treeNode)
        {
            // Conditions for the basic views of the subtree (content, control, and raw)
            // are available as fields of TreeWalker, and one of these is used in the
            // following code.
            TreeNode childTreeNode = treeNode.Nodes.Add(AEToString(rootElement));

            AutomationElement elementNode = TreeWalker.RawViewWalker.GetFirstChild(rootElement);
            while (elementNode != null)
            {
                WalkRawElements(elementNode, childTreeNode);
                elementNode = TreeWalker.RawViewWalker.GetNextSibling(elementNode);
            }
        }

        private static void WalkControlElements(AutomationElement rootElement, TreeNode treeNode)
        {
            // Conditions for the basic views of the subtree (content, control, and raw)
            // are available as fields of TreeWalker, and one of these is used in the
            // following code.
            AutomationElement elementNode = TreeWalker.ControlViewWalker.GetFirstChild(rootElement);

            while (elementNode != null)
            {
                TreeNode childTreeNode = treeNode.Nodes.Add(AEToString(elementNode));
                WalkControlElements(elementNode, childTreeNode);
                elementNode = TreeWalker.ControlViewWalker.GetNextSibling(elementNode);
            }
        }

        private static void WalkContentElements(AutomationElement rootElement, TreeNode treeNode)
        {
            // Conditions for the basic views of the subtree (content, control, and raw)
            // are available as fields of TreeWalker, and one of these is used in the
            // following code.
            AutomationElement elementNode = TreeWalker.ContentViewWalker.GetFirstChild(rootElement);

            while (elementNode != null)
            {
                TreeNode childTreeNode = treeNode.Nodes.Add(AEToString(elementNode));
                WalkContentElements(elementNode, childTreeNode);

                elementNode = TreeWalker.ContentViewWalker.GetNextSibling(elementNode);
            }
        }

        private static string TreeNodeToString(TreeNode treeNode)
        {
            StringBuilder sb = new StringBuilder();
            foreach (TreeNode node in treeNode.Nodes)
            {
                BuildTreeString(node, sb);
            }

            return sb.ToString();
        }

        private static void BuildTreeString(TreeNode parentNode, StringBuilder sb)
        {
            sb.Append(new string('\t', parentNode.Level));
            sb.Append(parentNode.Text);
            sb.Append(Environment.NewLine);
            foreach (TreeNode childNode in parentNode.Nodes)
            {
                BuildTreeString(childNode, sb);
            }
        }

        /*
        private static string WalkControlElementsToString(AutomationElement parentElement)
        {
            TreeNode rootNode = new TreeNode();
            WalkControlElements(parentElement, rootNode);
            return TreeNodeToString(rootNode);
        }

        private static string WalkContentElementsToString(AutomationElement parentElement)
        {
            TreeNode rootNode = new TreeNode();
            WalkContentElements(parentElement, rootNode);
            return TreeNodeToString(rootNode);
        }
        */
    }
}
