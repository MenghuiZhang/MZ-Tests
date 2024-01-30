# coding: utf8
import clr
from Autodesk.Revit.DB import Line, XYZ
from Autodesk.Revit.DB import SetComparisonResult, IntersectionResultArray,ClosestPointsPairBetweenTwoCurves
from IGF_Klasse import List

import sys
import clr
import os
sys.path.Add(os.path.dirname(__file__)+'\\External')
clr.AddReference('MouseKeyHook')
clr.AddReference('Gma.System.MouseKeyHook')
import MouseKeyHook.WindowsHelper as WindowsHelper
import MouseKeyHook.TemplateMethode as TemplateMethode
from Gma.System.MouseKeyHook import Hook


# public class Test
# {
#     private IKeyboardMouseEvents m_GlobalHook;
#     public void Subscribe()
#     {
#         // Note: for the application hook, use the Hook.AppEvents() instead
#         m_GlobalHook = Hook.GlobalEvents();

#         m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
#         m_GlobalHook.KeyPress += GlobalHookKeyPress;
#     }
#     private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
#     {
#         Console.WriteLine("KeyPress: \t{0}", e.KeyChar);
#     }
#     private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
#     {
#         Console.WriteLine("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp);

#         // uncommenting the following line will suppress the middle mouse button click
#         // if (e.Button == MouseButtons.Middle) { e.Handled = true; }
#     }

#     public void Unsubscribe()
#     {
#         m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
#         m_GlobalHook.KeyPress -= GlobalHookKeyPress;

#         //It is recommened to dispose it
#         m_GlobalHook.Dispose();
#     }

class HookTemplate:
    def __init__(self):
        self.Hook = None
