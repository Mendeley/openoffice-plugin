#!/usr/bin/env python

from __future__ import print_function
import os
import sys
import xml.sax.saxutils
import zipfile

def convert_vba_script_to_xml(input_path, module_name, debug_mode):
    xml_header = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:language="StarBasic">'''
    xml_footer = '</script:module>'
    vba_script = file(input_path).read()
    vba_script = vba_script.replace('${DEBUG_MODE}', debug_mode)
    xml_script = xml_header + xml.sax.saxutils.escape(vba_script) + xml_footer
    return xml_script

if __name__ == "__main__":
    plugin_version = sys.argv[1]
    debug_mode = sys.argv[2]

    print('LibreOffice plugin version: %s' % plugin_version)

    # Populate extension with files from the template dir
    EXTENSION_TEMPLATE_DIR = 'MendeleyEmptyExtension.oxt'
    EXTENSION_SOURCE_DIRS = ['external', 'src']

    extension_archive = zipfile.ZipFile('Mendeley-%s.oxt' % plugin_version, 'w')
    for dir_path, dirnames, filenames in os.walk(EXTENSION_TEMPLATE_DIR):
        for name in filenames:
            file_path = dir_path + '/' + name
            if name == 'description.xml':
                description_content = file(file_path).read().replace('%PLUGIN_VERSION%', plugin_version)
                extension_archive.writestr(name, description_content)
            else:
                extension_archive.write(file_path, os.path.relpath(file_path, EXTENSION_TEMPLATE_DIR))

    # Preprocess VBA source files and save to plugin archive
    for source_dir in EXTENSION_SOURCE_DIRS:
        for dir_path, dirnames, filenames in os.walk(source_dir):
            for name in filenames:
                if name.endswith('.vb'):
                    # VBA files are converted to XML files with a single <script> element
                    # and stored in Mendeley/
                    basename = os.path.splitext(name)[0]
                    module_name = basename[0].upper() + basename[1:]
                    xml_source = convert_vba_script_to_xml(dir_path + '/' + name, module_name, debug_mode)
                    extension_archive.writestr('Mendeley/%s.xba' % basename, xml_source)

    # Concatenate Python source files and save to plugin archive
    plugin_python_script = ''
    python_sources = ['src/MendeleyHttpClient.py', 'src/MendeleyDesktopAPI.py']
    for python_source_path in python_sources:
        plugin_python_script += file(python_source_path).read()
    extension_archive.writestr('Scripts/MendeleyDesktopAPI.py', plugin_python_script)

    extension_archive.close()
    print('Successfully built LibreOffice plugin version %s' % plugin_version)
