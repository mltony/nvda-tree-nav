# -*- coding: UTF-8 -*-
#A part of the TreeNav addon for NVDA
#Copyright (C) 2017-2024 Tony Malykh
#This file is covered by the GNU General Public License.
#See the file LICENSE  for more details.


import addonHandler
import api
from appModules.devenv import VsWpfTextViewTextInfo
from compoundDocuments import CompoundTextInfo
import controlTypes
import config
from NVDAObjects.IAccessible.chromium import ChromeVBufTextInfo

try:
    from config.configFlags import ReportLineIndentation
except (ImportError, ModuleNotFoundError):
    pass
import core
import ctypes
from enum import Enum, auto
import globalCommands
import globalPluginHandler
import gui
from gui.settingsDialogs import SettingsPanel
from gui import guiHelper, nvdaControls
import inputCore
import json
import keyboardHandler
from logHandler import log
import NVDAHelper
from NVDAObjects.IAccessible import IAccessible
from NVDAObjects.IAccessible import IA2TextTextInfo
from NVDAObjects.IAccessible.ia2TextMozilla import MozillaCompoundTextInfo
from NVDAObjects import NVDAObject, NVDAObjectTextInfo
import operator
import os
import queue
import re
import scriptHandler
from scriptHandler import script
import speech
import struct
import subprocess
import textInfos
from textInfos import UNIT_CHARACTER
from textInfos.offsets import OffsetsTextInfo
import threading
import time
import tones
from typing import Tuple
import ui
from utils.displayString import DisplayStringIntEnum
import versionInfo
import winUser
import wx
import dataclasses

try:
    ROLE_EDITABLETEXT = controlTypes.ROLE_EDITABLETEXT
    ROLE_TREEVIEWITEM = controlTypes.ROLE_TREEVIEWITEM
except AttributeError:
    ROLE_EDITABLETEXT = controlTypes.Role.EDITABLETEXT
    ROLE_TREEVIEWITEM = controlTypes.Role.TREEVIEWITEM

BUILD_YEAR = getattr(versionInfo, "version_year", 2023)

debug = False
if debug:
    LOG_FILE_NAME = r"H:\\2.txt"
    f = open(LOG_FILE_NAME, "w", encoding='utf=8')
    f.close()
    LOG_MUTEX = threading.Lock()
    def mylog(s):
        with LOG_MUTEX:
            f = open(LOG_FILE_NAME, "a", encoding='utf-8')
            print(s, file=f)
            f.close()
else:
    def mylog(*arg, **kwarg):
        pass


# Adapted from NVDA's speech module to count tabs as blank characters.
BLANK_CHUNK_CHARS = frozenset((" ", "\n", "\r", "\t", "\0", u"\xa0"))
def isBlank(text):
    return not text or set(text) <= BLANK_CHUNK_CHARS


def myAssert(condition):
    if not condition:
        raise RuntimeError("Assertion failed")

def initConfiguration():
    confspec = {
        "crackleVolume" : "integer( default=25, min=0, max=100)",
        "noNextTextChimeVolume" : "integer( default=50, min=0, max=100)",
        "noNextTextMessage" : "boolean( default=False)",
    }
    config.conf.spec["treeNav"] = confspec


def getConfig(key):
    value = config.conf["treeNav"][key]
    return value

def setConfig(key, value):
    config.conf["treeNav"][key] = value


addonHandler.initTranslation()
initConfiguration()

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = _("TreeNav")
    def __init__(self, *args, **kwargs):
        super(GlobalPlugin, self).__init__(*args, **kwargs)

    def terminate(self):
        pass

    def chooseNVDAObjectOverlayClasses (self, obj, clsList):
        if obj.role == ROLE_TREEVIEWITEM:
            clsList.append(TreeIndentNav)
            return

class Beeper:
    BASE_FREQ = speech.IDT_BASE_FREQUENCY
    def getPitch(self, indent):
        return self.BASE_FREQ*2**(indent/24.0) #24 quarter tones per octave.

    BEEP_LEN = 10 # millis
    PAUSE_LEN = 5 # millis
    MAX_CRACKLE_LEN = 400 # millis
    MAX_BEEP_COUNT = MAX_CRACKLE_LEN // (BEEP_LEN + PAUSE_LEN)


    def fancyCrackle(self, levels, volume):
        levels = self.uniformSample(levels, self.MAX_BEEP_COUNT )
        beepLen = self.BEEP_LEN
        pauseLen = self.PAUSE_LEN
        pauseBufSize = NVDAHelper.generateBeep(None,self.BASE_FREQ,pauseLen,0, 0)
        beepBufSizes = [NVDAHelper.generateBeep(None,self.getPitch(l), beepLen, volume, volume) for l in levels]
        bufSize = sum(beepBufSizes) + len(levels) * pauseBufSize
        buf = ctypes.create_string_buffer(bufSize)
        bufPtr = 0
        for l in levels:
            bufPtr += NVDAHelper.generateBeep(
                ctypes.cast(ctypes.byref(buf, bufPtr), ctypes.POINTER(ctypes.c_char)),
                self.getPitch(l), beepLen, volume, volume)
            bufPtr += pauseBufSize # add a short pause
        tones.player.stop()
        tones.player.feed(buf.raw)

    def simpleCrackle(self, n, volume):
        return self.fancyCrackle([0] * n, volume)


    NOTES = "A,B,H,C,C#,D,D#,E,F,F#,G,G#".split(",")
    NOTE_RE = re.compile("[A-H][#]?")
    BASE_FREQ = 220
    def getChordFrequencies(self, chord):
        myAssert(len(self.NOTES) == 12)
        prev = -1
        result = []
        for m in self.NOTE_RE.finditer(chord):
            s = m.group()
            i =self.NOTES.index(s)
            while i < prev:
                i += 12
            result.append(int(self.BASE_FREQ * (2 ** (i / 12.0))))
            prev = i
        return result

    def fancyBeep(self, chord, length, left=10, right=10):
        beepLen = length
        freqs = self.getChordFrequencies(chord)
        intSize = 8 # bytes
        bufSize = max([NVDAHelper.generateBeep(None,freq, beepLen, right, left) for freq in freqs])
        if bufSize % intSize != 0:
            bufSize += intSize
            bufSize -= (bufSize % intSize)
        tones.player.stop()
        bbs = []
        result = [0] * (bufSize//intSize)
        for freq in freqs:
            buf = ctypes.create_string_buffer(bufSize)
            NVDAHelper.generateBeep(buf, freq, beepLen, right, left)
            bytes = bytearray(buf)
            unpacked = struct.unpack("<%dQ" % (bufSize // intSize), bytes)
            result = map(operator.add, result, unpacked)
        maxInt = 1 << (8 * intSize)
        result = map(lambda x : x %maxInt, result)
        packed = struct.pack("<%dQ" % (bufSize // intSize), *result)
        tones.player.feed(packed)

    def uniformSample(self, a, m):
        n = len(a)
        if n <= m:
            return a
        # Here assume n > m
        result = []
        for i in range(0, m*n, n):
            result.append(a[i // m])
        return result


class TreeIndentNav(NVDAObject):
    scriptCategory = _("TreeNav")
    beeper = Beeper()

    @script(description=_("Moves to the next item on the same level within current subtree."), gestures=['kb:NVDA+alt+DownArrow'])
    def script_moveToNextSibling(self, gesture):
        # Translators: error message if next sibling couldn't be found in Tree view
        errorMsg = _("No next item on the same level within this subtree")
        self.moveInTree(1, errorMsg, op=operator.eq)

    @script(description=_("Moves to the previous item on the same level within current subtree."), gestures=['kb:NVDA+alt+UpArrow'])
    def script_moveToPreviousSibling(self, gesture):
        # Translators: error message if next sibling couldn't be found in Tree view
        errorMsg = _("No previous item on the same level within this subtree")
        self.moveInTree(-1, errorMsg, op=operator.eq)

    @script(description=_("Moves to the next item on the same level."), gestures=['kb:NVDA+Control+alt+DownArrow'])
    def script_moveToNextSiblingForce(self, gesture):
        # Translators: error message if next sibling couldn't be found in Tree view
        errorMsg = _("No next item on the same level in this tree view")
        self.moveInTree(1, errorMsg, op=operator.eq, unbounded=True)

    @script(description=_("Moves to the previous item on the same level."), gestures=['kb:NVDA+Control+alt+UpArrow'])
    def script_moveToPreviousSiblingForce(self, gesture):
        # Translators: error message if previous sibling couldn't be found in Tree view
        errorMsg = _("No previous item on the same level in this tree view")
        self.moveInTree(-1, errorMsg, op=operator.eq, unbounded=True)

    @script(description=_("Moves to the last item on the same level within current subtree."), gestures=['kb:NVDA+alt+Shift+DownArrow'])
    def script_moveToLastSibling(self, gesture):
        # Translators: error message if next sibling couldn't be found in Tree view
        errorMsg = _("No next item on the same level within this subtree")
        self.moveInTree(1, errorMsg, op=operator.eq, moveCount=1000)

    @script(description=_("Moves to the first item on the same level within current subtree."), gestures=['kb:NVDA+alt+Shift+UpArrow'])
    def script_moveToFirstSibling(self, gesture):
        # Translators: error message if next sibling couldn't be found in Tree view
        errorMsg = _("No previous item on the same level within this subtree")
        self.moveInTree(-1, errorMsg, op=operator.eq, moveCount=1000)

    @script(description=_("Speak parent item."), gestures=['kb:NVDA+I'])
    def script_speakParent(self, gesture):
        count=scriptHandler.getLastScriptRepeatCount()
        # Translators: error message if parent couldn't be found)
        errorMsg = _("No parent item in this tree view")
        self.moveInTree(-1, errorMsg, unbounded=True, op=operator.lt, speakOnly=True, moveCount=count+1)

    @script(description=_("Moves to the next child in tree view."), gestures=['kb:NVDA+alt+RightArrow'])
    def script_moveToChild(self, gesture):
        # Translators: error message if a child couldn't be found
        errorMsg = _("NO child")
        self.moveInTree(1, errorMsg, unbounded=False, op=operator.gt)

    @script(description=_("Moves to parent in tree view."), gestures=['kb:NVDA+alt+LeftArrow'])
    def script_moveToParent(self, gesture):
        # Translators: error message if parent couldn't be found
        errorMsg = _("No parent")
        self.moveInTree(-1, errorMsg, unbounded=True, op=operator.lt)

    def getLevel(self, obj):
        try:
            return obj.positionInfo["level"]
        except AttributeError:
            return None
        except KeyError:
            return None


    def moveInTree(self, increment, errorMessage, unbounded=False, op=operator.eq, speakOnly=False, moveCount=1):
        obj = api.getFocusObject()
        level = self.getLevel(obj)
        found = False
        levels = []
        while True:
            if increment > 0:
                obj = obj.next
            else:
                obj = obj.previous
            newLevel = self.getLevel(obj)
            if newLevel is None:
                break
            if op(newLevel, level):
                found = True
                level = newLevel
                result = obj
                moveCount -= 1
                if moveCount == 0:
                    break
            elif newLevel < level:
                # Not found in this subtree
                if not unbounded:
                    break
            levels.append(newLevel )

        if found:
            self.beeper.fancyCrackle(levels, volume=getConfig("crackleVolume"))
            if not speakOnly:
                result.setFocus()
            else:
                speech.speakObject(result)
        else:
            self.endOfDocument(errorMessage)

    def endOfDocument(self, message=None):
        volume = getConfig("noNextTextChimeVolume")
        self.beeper.fancyBeep("HF", 100, volume, volume)
        if getConfig("noNextTextMessage") and message is not None:
            ui.message(message)
