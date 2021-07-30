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

import os
import traceback

import uno
import unohelper
from com.sun.star.awt import Selection
from com.sun.star.awt.Key import RETURN as KEY_RETURN
from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD
from com.sun.star.beans import PropertyValue
from com.sun.star.drawing.FillStyle import NONE as FS_NONE, SOLID as FS_SOLID
from com.sun.star.lang import Locale
from com.sun.star.task import XJobExecutor

import pygments.util
from pygments import styles
from pygments.lexers import get_all_lexers
from pygments.lexers import get_lexer_by_name
from pygments.lexers import guess_lexer
from pygments.styles import get_all_styles


class CodeHighlighter(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        try:
            self.ctx = ctx
            self.sm = ctx.ServiceManager
            desktop = self.create("com.sun.star.frame.Desktop")
            self.doc = desktop.getCurrentComponent()
            self.cfg_access = self.create_cfg_access()
            self.config = self.load_config()
            self.dialog = self.create_dialog()
        except Exception:
            traceback.print_exc()

    # XJobExecutor
    def trigger(self, arg):
        # print(f"trigger arg: {arg}")
        try:
            getattr(self, 'do_'+arg)()
        except:
            traceback.print_exc()

    # core functions
    def create(self, service):
        return self.sm.createInstance(service)

    def to_int(self, hex_str):
        if hex_str:
            return int(hex_str[-6:], 16)
        return 0

    def create_cfg_access(self):
        cfg = self.create('com.sun.star.configuration.ConfigurationProvider')
        prop = PropertyValue('nodepath', 0, '/ooo.ext.code-highlighter.Registry/Settings', 0)
        cfg_access = cfg.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess', (prop,))
        return cfg_access

    def load_config(self):
        properties = self.cfg_access.ElementNames
        values = self.cfg_access.getPropertyValues(properties)
        return dict(zip(properties, values))

    def save_config(self, **kwargs):
        self.config.update(kwargs)
        self.cfg_access.setPropertyValues(tuple(kwargs.keys()), tuple(kwargs.values()))
        self.cfg_access.commitChanges()

    def create_dialog(self):
        # get_all_lexers() returns:
        # (longname, tuple of aliases, tuple of filename patterns, tuple of mimetypes)
        all_lexers = sorted((lex[0] for lex in get_all_lexers()), key=str.casefold)
        all_lexer_aliases = [lex[0] for lex in get_all_lexers()]
        for lex in get_all_lexers():
            all_lexer_aliases.extend(list(lex[1]))
        all_styles = sorted(get_all_styles(), key=lambda x: (x != 'default', x.lower()))

        dialog_provider = self.create("com.sun.star.awt.DialogProvider")
        dialog = dialog_provider.createDialog("vnd.sun.star.extension://javahelps.codehighlighter/dialogs/CodeHighlighter.xdl")

        cb_lang = dialog.getControl('cb_lang')
        cb_style = dialog.getControl('cb_style')
        check_col_bg = dialog.getControl('check_col_bg')

        # TODO: reformat config access
        cb_lang.addItem('automatic', 0)
        cb_lang.Text = self.config['Language']
        cb_lang.setSelection(Selection(0, len(cb_lang.Text)))
        cb_lang.addItems(all_lexers, 0)

        style = self.config['Style']
        if style in all_styles:
            cb_style.Text = style
        cb_style.addItems(all_styles, 0)

        check_col_bg.State = int(self.config['ColorizeBackground'])

        return dialog

    def do_highlight(self):
        # self.dialog.setVisible(True)
        # 0: canceled, 1: OK
        if self.dialog.execute() == 0:
            return

        lang = self.dialog.getControl('cb_lang').Text
        style = self.dialog.getControl('cb_style').Text
        colorize_bg = self.dialog.getControl('check_col_bg').State
        if lang == 'automatic':
            lang = None
        # # TODO: handle exceptions here
        # assert lang == None or (lang in all_lexer_aliases), 'no valid language: ' + lang
        # assert style in all_styles, 'no valid style: ' + style

        self.save_config(Style=style, Language=lang or 'automatic', ColorizeBackground=str(colorize_bg))

        self.highlight_source_code(lang, style, colorize_bg != 0)

    def do_highlight_previous(self):
        lang = self.config['Language']
        style = self.config['Style']
        colorize_bg = int(self.config['ColorizeBackground'])
        if lang == 'automatic':
            lang = None
        self.highlight_source_code(lang, style, colorize_bg != 0)

    def highlight_source_code(self, lang, style_, colorize_bg=False):
        style = styles.get_style_by_name(style_)
        style.name = style_
        bg_color = style.background_color if colorize_bg else None

        # Get the selected item
        selected_item = self.doc.getCurrentController().getSelection()
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
                        code_block.FillColor = self.to_int(bg_color)
                    code = code_block.String
                    cursor = code_block.createTextCursor()
                    cursor.gotoStart(False)
                else:
                    # Plain text
                    # highlight_code_string(style, lang, code_block)
                    code_block.CharLocale = Locale("zxx", "", "")
                    code_block.ParaBackColor = -1
                    if bg_color:
                        code_block.ParaBackColor = self.to_int(bg_color)
                    code = code_block.getString()
                    cursor = code_block.getText().createTextCursorByRange(code_block)
                    cursor.goLeft(0, False)
                self.highlight_code(code, cursor, lang, style)

        elif selected_item.supportsService('com.sun.star.text.TextFrame'):
            # Selection is a text frame
            code_block = selected_item
            code_block.BackColor = -1
            if bg_color:
                code_block.BackColor = self.to_int(bg_color)
            code = code_block.String
            cursor = code_block.createTextCursorByRange(code_block)
            cursor.CharLocale = Locale("zxx", "", "")
            cursor.gotoStart(False)
            self.highlight_code(code, cursor, lang, style)

        elif selected_item.supportsService('com.sun.star.text.TextTableCursor'):
            # Selection is one or more table cell range
            selected_item.CharLocale = Locale("zxx", "", "")
            table = self.doc.CurrentController.ViewCursor.TextTable
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
                            code_block.BackColor = self.to_int(bg_color)
                        code = code_block.String
                        cursor = code_block.createTextCursor()
                        cursor.gotoStart(False)
                        self.highlight_code(code, cursor, lang, style)
            else:
                # only one cell
                code_block = table.getCellByName(rangename)
                code_block.BackColor = -1
                if bg_color:
                    code_block.BackColor = self.to_int(bg_color)
                code = code_block.String
                cursor = code_block.createTextCursor()
                cursor.gotoStart(False)
                self.highlight_code(code, cursor, lang, style)

        elif hasattr(selected_item, 'SupportedServiceNames') and selected_item.supportsService('com.sun.star.text.TextCursor'):
            # LO Impress text selection
            code_block = selected_item
            code = code_block.getString()
            cursor = code_block.getText().createTextCursorByRange(code_block)
            cursor.goLeft(0, False)
            self.highlight_code(code, cursor, lang, style)

    def highlight_code(self, code, cursor, lang, style):
        if lang is None:
            lexer = guess_lexer(code)
            # print(f'lexer name = {lexer.name}')
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

        self.doc.lockControllers()
        undomanager = self.doc.UndoManager
        undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {style.name})")
        try:
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
                            cursor.CharColor = self.to_int(tok_style['color'])
                            cursor.CharWeight = W_BOLD if tok_style['bold'] else W_NORMAL
                            cursor.CharPosture = SL_ITALIC if tok_style['italic'] else SL_NONE
                        except:
                            pass
                        finally:
                            cursor.goRight(0, False)  # deselects the selected text
                    lastval = tok_value
                    lasttype = tok_type
        except Exception:
            traceback.print_exc()
        finally:
            undomanager.leaveUndoContext()
            self.doc.unlockControllers()


# Component registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(CodeHighlighter, "code.highlighter.impl", (),)
