# coding: utf8
from pyrevit.api import AdWindows
from IGF_MauseKeyHook import WindowsHelper
def Nummer_Korrigieren(nummer,stellen):
    if len(str(nummer)) < stellen:
        return (stellen-len(str(nummer))) * '0' + str(nummer)
    return str(nummer)


def RVTMainWindowActive():
    childHwnd = WindowsHelper.SendMessage(AdWindows.ComponentManager.ApplicationWindow, 0x0229, 0, 0)
    WindowsHelper.SendMessage(AdWindows.ComponentManager.ApplicationWindow,0x0222, childHwnd, 0)
