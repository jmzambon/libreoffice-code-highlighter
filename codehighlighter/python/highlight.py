# Code Highligher is a LibreOffice extension to highlight code snippets
# over 350 languages.

# Copyright (C) 2017  Gobinath

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import uno
from com.sun.star.awt import Selection
from com.sun.star.awt.Key import RETURN as KEY_RETURN
from com.sun.star.drawing.FillStyle import NONE as FS_NONE, SOLID as FS_SOLID
from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD

from com.sun.star.beans import PropertyValue
from com.sun.star.lang import Locale

from pygments import styles
from pygments.lexers import get_all_lexers
from pygments.lexers import get_lexer_by_name
from pygments.lexers import guess_lexer
from pygments.styles import get_all_styles
import pygments.util
import os


def rgb(r, g, b):
    return (r & 255) << 16 | (g & 255) << 8 | (b & 255)


def to_int(hex_str):
    if hex_str:
        return int(hex_str[-6:], 16)
    return 0


def log(msg):
    with open("/tmp/code-highlighter.log", "a") as text_file:
        text_file.write(str(msg) + "\r\n\r\n")


def create_dialog():
    # get_all_lexers() returns:
    # (longname, tuple of aliases, tuple of filename patterns, tuple of mimetypes)
    all_lexers = sorted((lex[0] for lex in get_all_lexers()), key=str.casefold)
    all_lexer_aliases = [lex[0] for lex in get_all_lexers()]
    for lex in get_all_lexers():
        all_lexer_aliases.extend(list(lex[1]))
    all_styles = sorted(get_all_styles(), key=lambda x: (x != 'default', x.lower()))

    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    dialog_provider = smgr.createInstance("com.sun.star.awt.DialogProvider")
    dialog = dialog_provider.createDialog("vnd.sun.star.extension://javahelps.codehighlighter/dialogs/CodeHighlighter.xdl")

    cfg = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
    prop = PropertyValue()
    prop.Name = 'nodepath'
    prop.Value = '/ooo.ext.code-highlighter.Registry/Settings'
    cfg_access = cfg.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess', (prop,))

    cb_lang = dialog.getControl('cb_lang')
    cb_style = dialog.getControl('cb_style')
    check_col_bg = dialog.getControl('check_col_bg')

    cb_lang.addItem('automatic', 0)
    cb_lang.Text = cfg_access.getPropertyValue('Language')
    cb_lang.setSelection(Selection(0, len(cb_lang.Text)))
    cb_lang.addItems(all_lexers, 0)

    style = cfg_access.getPropertyValue('Style')
    if style in all_styles:
        cb_style.Text = style
    cb_style.addItems(all_styles, 0)

    check_col_bg.State = int(cfg_access.getPropertyValue('ColorizeBackground'))

    dialog.setVisible(True)
    # 0: canceled, 1: OK
    if dialog.execute() == 0:
        return

    lang = cb_lang.Text
    style = cb_style.Text
    colorize_bg = check_col_bg.State
    if lang == 'automatic':
        lang = None
    assert lang == None or (lang in all_lexer_aliases), 'no valid language: ' + lang
    assert style in all_styles, 'no valid style: ' + style

    cfg_access.setPropertyValue('Style', style)
    cfg_access.setPropertyValue('Language', lang or 'automatic')
    cfg_access.setPropertyValue('ColorizeBackground', str(colorize_bg))
    cfg_access.commitChanges()

    highlightSourceCode(lang, style, colorize_bg != 0)


def apply_previous_settings():
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    cfg = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
    prop = PropertyValue()
    prop.Name = 'nodepath'
    prop.Value = '/ooo.ext.code-highlighter.Registry/Settings'
    cfg_access = cfg.createInstanceWithArguments('com.sun.star.configuration.ConfigurationAccess', (prop,))

    lang = cfg_access.getPropertyValue('Language')
    style = cfg_access.getPropertyValue('Style')
    colorize_bg = int(cfg_access.getPropertyValue('ColorizeBackground'))

    if lang == 'automatic':
        lang = None

    highlightSourceCode(lang, style, colorize_bg != 0)


def highlightSourceCode(lang, style_, colorize_bg=False):
    style = styles.get_style_by_name(style_)
    bg_color = style.background_color if colorize_bg else None

    ctx = XSCRIPTCONTEXT
    doc = ctx.getDocument()
    doc.lockControllers()
    undomanager = doc.UndoManager
    undomanager.enterUndoContext(f"code highlight (lang: {lang or 'automatic'}, style: {style_})")
    try:
        # Get the selected item
        selected_item = doc.getCurrentController().getSelection()
        if hasattr(selected_item, 'getCount'):
            for item_idx in range(selected_item.getCount()):
                code_block = selected_item.getByIndex(item_idx)
                if code_block.supportsService('com.sun.star.drawing.Text'):
                    # TextBox
                    # highlight_code(style, lang, code_block)
                    code_block.CharLocale = Locale("zxx", "", "")
                    code_block.FillStyle = FS_NONE
                    if bg_color:
                        code_block.FillStyle = FS_SOLID
                        code_block.FillColor = to_int(bg_color)
                    code = code_block.String
                    cursor = code_block.createTextCursor()
                    cursor.gotoStart(False)
                else:
                    # Plain text
                    # highlight_code_string(style, lang, code_block)
                    code_block.CharLocale = Locale("zxx", "", "")
                    code_block.ParaBackColor = -1
                    if bg_color:
                        code_block.ParaBackColor = to_int(bg_color)
                    code = code_block.getString()
                    cursor = code_block.getText().createTextCursorByRange(code_block)
                    cursor.goLeft(0, False)
                highlight_code(code, cursor, lang, style)

        elif selected_item.supportsService('com.sun.star.text.TextFrame'):
            # Selection is a text frame
            code_block = selected_item
            code_block.BackColor = -1
            if bg_color:
                code_block.BackColor = to_int(bg_color)
            code = code_block.String
            cursor = code_block.createTextCursorByRange(code_block)
            cursor.CharLocale = Locale("zxx", "", "")
            cursor.gotoStart(False)
            highlight_code(code, cursor, lang, style)

        elif selected_item.supportsService('com.sun.star.text.TextTableCursor'):
            # Selection is one or more table cell range
            selected_item.CharLocale = Locale("zxx", "", "")
            table = doc.CurrentController.ViewCursor.TextTable
            rangename = selected_item.RangeName
            if ':' in rangename:
                # at least two cells
                cellrange = table.getCellRangeByName(rangename)
                nrows, ncols = len(cellrange.Data), len(cellrange.Data[0])
                for row in range(nrows):
                    for col in range(ncols):
                        code_block = cellrange.getCellByPosition(col, row)
                        code_block.BackColor = -1
                        if bg_color:
                            code_block.BackColor = to_int(bg_color)
                        code = code_block.String
                        cursor = code_block.createTextCursor()
                        cursor.gotoStart(False)
                        highlight_code(code, cursor, lang, style)
            else:
                # only one cell
                code_block = table.getCellByName(rangename)
                code_block.BackColor = -1
                if bg_color:
                    code_block.BackColor = to_int(bg_color)
                code = code_block.String
                cursor = code_block.createTextCursor()
                cursor.gotoStart(False)
                highlight_code(code, cursor, lang, style)

        elif hasattr(selected_item, 'SupportedServiceNames') and selected_item.supportsService('com.sun.star.text.TextCursor'):
            # LO Impress text selection
            code_block = selected_item
            code = code_block.getString()
            cursor = code_block.getText().createTextCursorByRange(code_block)
            cursor.goLeft(0, False)
            highlight_code(code, cursor, lang, style)
    finally:
        undomanager.leaveUndoContext()
        doc.unlockControllers()


def highlight_code(code, cursor, lang, style):
    if lang is None:
        lexer = guess_lexer(code)
    else:
        try:
            lexer = get_lexer_by_name(lang)
        except pygments.util.ClassNotFound:
            # get_lexer_by_name() only checks aliases, not the actual longname
            for lex in get_all_lexers():
                if lex[0] == lang:
                    # found the longname, use the first alias
                    lexer = get_lexer_by_name(lex[1][0])
                    break
            else:
                raise
    # prevent offset color if selection start with empty line
    lexer.stripnl = False
    # caching consecutive tokens with same token type
    lastval = ''
    lasttype = None
    for tok_type, tok_value in lexer.get_tokens(code):
        if tok_type == lasttype:
            lastval += tok_value
        else:
            if lastval:
                cursor.goRight(len(lastval), True)  # selects the token's text
                try:
                    tok_style = style.style_for_token(lasttype)
                    cursor.CharColor = to_int(tok_style['color'])
                    cursor.CharWeight = W_BOLD if tok_style['bold'] else W_NORMAL
                    cursor.CharPosture = SL_ITALIC if tok_style['italic'] else SL_NONE
                except:
                    pass
                finally:
                    cursor.goRight(0, False)  # deselects the selected text
            lastval = tok_value
            lasttype = tok_type


# export these, so they can be used for key bindings
g_exportedScripts = (create_dialog, apply_previous_settings)
