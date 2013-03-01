#!/usr/bin/env python

import os, os.path

# Generates resource file from each po file.
# And also other configuration stuff too.

desc_h = """<?xml version='1.0' encoding='UTF-8'?>
<description xmlns="http://openoffice.org/extensions/description/2006"
xmlns:xlink="http://www.w3.org/1999/xlink"
xmlns:d="http://openoffice.org/extensions/description/2006">
<identifier value="mytools.calc.WatchWindow" />
<version value="{VERSION}" />
<dependencies>
<OpenOffice.org-minimal-version value="3.4" d:name="OpenOffice.org 3.4" />
</dependencies>
<registration>
<simple-license accept-by="admin" default-license-id="this" suppress-on-update="true" suppress-if-required="true">
<license-text xlink:href="LICENSE" lang="en" license-id="this" />
</simple-license>
</registration>
<display-name>
{NAMES}
</display-name>
<extension-description>
{DESCRIPTIONS}
</extension-description>
<update-information>
<src xlink:href="https://raw.github.com/hanya/WatchingWindow/master/files/WatchingWindow.update.xml"/>
</update-information>
</description>"""

update_feed = """<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/update/2006" 
xmlns:xlink="http://www.w3.org/1999/xlink"
xmlns:d="http://openoffice.org/extensions/description/2006">
<identifier value="mytools.calc.WatchWindow" />
<version value="{VERSION}" />
<dependencies>
<d:OpenOffice.org-minimal-version value="3.4" d:name="OpenOffice.org 3.4" />
</dependencies>
<update-download>
<src xlink:href="https://raw.github.com/hanya/WatchingWindow/master/files/WatchingWindow-{VERSION}.oxt"/>
</update-download>
</description>
"""


def genereate_description(d):
    version = read_version()
    
    names = []
    for lang, v in d.iteritems():
        name = v["id.label.ww"]
        names.append("<name lang=\"{LANG}\">{NAME}</name>".format(LANG=lang, NAME=name.encode("utf-8")))
    
    descs = []
    for lang, v in d.iteritems():
        desc = v["id.extension.description"]
        with open("descriptions/desc_{LANG}.txt".format(LANG=lang), "w") as f:
            f.write(desc.encode("utf-8"))
        descs.append("<src lang=\"{LANG}\" xlink:href=\"descriptions/desc_{LANG}.txt\"/>".format(LANG=lang))
    
    return desc_h.format(
        VERSION=version, NAMES="\n".join(names), DESCRIPTIONS="\n".join(descs))


def read_version():
    version = ""
    with open("VERSION") as f:
        version = f.read().strip()
    return version


config_h = """<?xml version='1.0' encoding='UTF-8'?>
<oor:component-data 
  xmlns:oor="http://openoffice.org/2001/registry" 
  xmlns:xs="http://www.w3.org/2001/XMLSchema" 
  oor:package="{PACKAGE}" 
  oor:name="{NAME}">"""
config_f = "</oor:component-data>"


class XCUData(object):
    
    PACKAGE = ""
    NAME = ""
    
    def __init__(self):
        self.lines = []
    
    def append(self, line):
        self.lines.append(line)
    
    def add_node(self, name, op=None):
        if op:
            self.append("<node oor:name=\"{NAME}\" oor:op=\"{OP}\">".format(NAME=name, OP=op))
        else:
            self.append("<node oor:name=\"{NAME}\">".format(NAME=name))
    
    def close_node(self):
        self.append("</node>")
    
    def add_prop(self, name, value):
        self.append("<prop oor:name=\"{NAME}\">".format(NAME=name))
        self.append("<value>{VALUE}</value>".format(VALUE=value))
        self.append("</prop>")
    
    def open_prop(self, name):
        self.append("<prop oor:name=\"{NAME}\">".format(NAME=name))
    
    def close_prop(self):
        self.append("</prop>")
    
    def add_value(self, v, locale=None):
        if locale:
            self.append("<value xml:lang=\"{LANG}\">{VALUE}</value>".format(VALUE=v.encode("utf-8"), LANG=locale))
        else:
            self.append("<value>{VALUE}</value>".format(VALUE=v.encode("utf-8")))
    
    def add_value_for_localse(self, name, k, d):
        self.open_prop(name)
        locales = list(d.iterkeys())
        locales.sort()
        for lang in locales:
            _d = d[lang]
            self.add_value(_d[k], lang)
        self.close_prop()
    
    #def _generate(self, d): pass
    
    def generate(self, d):
        self.lines.append(config_h.format(PACKAGE=self.PACKAGE, NAME=self.NAME))
        self._generate(d)
        self.lines.append(config_f)
        return "\n".join(self.lines)


class CalcWindowStateXCU(XCUData):
    
    PACKAGE = "org.openoffice.Office.UI"
    NAME = "CalcWindowState"
    
    def _generate(self, d):
        
        self.add_node("UIElements")
        self.add_node("States")
        self.add_node("private:resource/toolpanel/mytools.calc/WatchWindow", "replace")
        
        self.add_value_for_localse("UIName", "id.label.ww", d)
        self.add_prop("ImageURL", "vnd.sun.star.extension://mytools.calc.WatchWindow/icons/ww_24.png")
        
        self.close_node()
        self.close_node()
        self.close_node()


def extract(d, locale, lines):
    msgid = msgstr = id = ""
    for l in lines:
        #if l[0] == "#":
        #    pass
        if l[0:2] == "#,":
            pass
        elif l[0:2] == "#:":
            id = l[2:].strip()
        if l[0] == "#":
            continue
        elif l.startswith("msgid"):
            msgid = l[5:]
        elif l.startswith("msgstr"):
            msgstr = l[6:].strip()
            #print(id, msgid, msgstr)
            if msgstr and id:
                d[id] = msgstr[1:-1].decode("utf-8").replace('\\"', '"')
        _l = l.strip()
        if not _l:
            continue


def as_resource(d):
    lines = []
    
    for k, v in d.iteritems():
        cs = []
        for c in v:
            a = ord(c)
            if a > 0x7f:
                cs.append("\\u%04x" % a)
            else:
                cs.append(c)
        lines.append("%s=%s" % (k, "".join(cs)))
    lines.sort()
    return "\n".join(lines)


def write_resource(res_path, d):
    lines = as_resource(d)
    with open(res_path, "w") as f:
        f.write("# comment\n")
        f.write(lines.encode("utf-8"))

def write_update_feed():
    version = read_version()
    s = update_feed.format(VERSION=version)
    with open("./files/WatchingWindow.update.xml", "w") as f:
        f.write(s.encode("utf-8"))

def main():
    prefix = "strings_"
    res_dir = "resources"
    
    locales = {}
    
    po_dir = os.path.join(".", "po")
    for po in os.listdir(po_dir):
        if po.endswith(".po"):
            locale = po[:-3]
            try:
                lines = open(os.path.join(po_dir, po)).readlines()
            except:
                print("%s can not be opened")
            d = {}
            extract(d, locale, lines)
            locales[locale] = d
    
    resources_dir = os.path.join(".", res_dir)
    
    for locale, d in locales.iteritems():
        write_resource(os.path.join(resources_dir, 
            "%s%s.properties" % (prefix, locale.replace("-", "_"))), d)
    
    s = CalcWindowStateXCU().generate(locales)
    with open("CalcWindowState.xcu", "w") as f:
        f.write(s)#.encode("utf-8"))
    
    s = genereate_description(locales)
    with open("description.xml", "w") as f:
        f.write(s)#.encode("utf-8"))
    
    write_update_feed()


if __name__ == "__main__":
    main()
