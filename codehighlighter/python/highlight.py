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

import traceback
from os import linesep as os_linesep

import unohelper
from com.sun.star.awt import Selection
from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, ERRORBOX
from com.sun.star.beans import PropertyValue
from com.sun.star.drawing.FillStyle import NONE as FS_NONE, SOLID as FS_SOLID
from com.sun.star.lang import Locale
from com.sun.star.sheet.CellFlags import STRING as CF_STRING
from com.sun.star.task import XJobExecutor
from com.sun.star.document import XUndoAction

import pygments
from pygments import styles
from pygments.lexers import get_all_lexers
from pygments.lexers import get_lexer_by_name
from pygments.lexers import guess_lexer
from pygments.styles import get_all_styles


class UndoAction(unohelper.Base, XUndoAction):
    """ Add undo/redo action for highlighting not catched by the system,
        i.e. when applied on textbox objects."""

    def __init__(self, doc, textbox, title):
        self.doc = doc
        self.textbox = textbox
        self.Title = title
        self.eol = len(os_linesep)
        self.old_portions = None
        self.old_bg = None
        self.new_portions = None
        self.new_bg = None
        self.charprops = ("CharColor", "CharLocale", "CharPosture", "CharWeight")
        self.bgprops = ("FillColor", "FillStyle")
        self.get_old_state()

    def undo(self):
        self._format(self.old_portions, self.old_bg)

    def redo(self):
        self._format(self.new_portions, self.new_bg)

    def get_old_state(self):
        self.old_bg = self.textbox.getPropertyValues(self.bgprops)
        self.old_portions = self._extract_portions()

    def get_new_state(self):
        self.new_bg = self.textbox.getPropertyValues(self.bgprops)
        self.new_portions = self._extract_portions()

    def _extract_portions(self):
        textportions = []
        for para in self.textbox:
            if textportions:    # new paragraph after first one
                textportions[-1][0] += self.eol
            for portion in para:
                plen = len(portion.String)
                pprops = portion.getPropertyValues(self.charprops)
                if textportions and textportions[-1][1] == pprops:
                    textportions[-1][0] += plen
                else:
                    textportions.append([plen, pprops])
        return textportions

    def _format(self, portions, bg):
        self.textbox.setPropertyValues(self.bgprops, bg)
        cursor = self.textbox.createTextCursor()
        cursor.gotoStart(False)
        for length, props in portions:
            cursor.goRight(length, True)
            cursor.setPropertyValues(self.charprops, props)
            cursor.collapseToEnd()
        self.doc.CurrentController.select(self.textbox)
        self.doc.setModified(True)


class CodeHighlighter(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        try:
            self.ctx = ctx
            self.sm = ctx.ServiceManager
            self.desktop = self.create("com.sun.star.frame.Desktop")
            self.doc = self.desktop.getCurrentComponent()
            self.frame = self.doc.CurrentController.Frame
            self.dispatcher = self.create("com.sun.star.frame.DispatchHelper")
            self.cfg_access = self.create_cfg_access()
            self.options = self.load_options()
            self.dialog = self.create_dialog()
            self.nolocale = Locale("zxx", "", "")
        except Exception:
            traceback.print_exc()

    # XJobExecutor
    def trigger(self, arg):
        # print(f"trigger arg: {arg}")
        try:
            getattr(self, 'do_'+arg)()
        except Exception:
            traceback.print_exc()

    # main functions
    def do_highlight(self):
        # get options choice
        # 0: canceled, 1: OK
        # self.dialog.setVisible(True)
        if self.dialog.execute() == 0:
            return
        lang = self.dialog.getControl('cb_lang').Text.strip() or 'automatic'
        style = self.dialog.getControl('cb_style').Text.strip() or 'default'
        colorize_bg = self.dialog.getControl('check_col_bg').State

        if lang != 'automatic' and lang not in self.all_lexer_aliases:
            self.msgbox("Unsupported language.")
            return
        if style not in self.all_styles:
            self.msgbox("Unknown.")
            return
        self.save_options(Style=style, Language=lang, ColorizeBackground=str(colorize_bg))
        self.highlight_source_code()

    def do_highlight_previous(self):
        self.highlight_source_code()

    # private functions
    def create(self, service):
        return self.sm.createInstance(service)

    def msgbox(self, message, boxtype=ERRORBOX, title="Error"):
        frame = self.desktop.ActiveFrame
        if frame.ActiveFrame:
            # top window is a subdocument
            frame = frame.ActiveFrame
        win = frame.ContainerWindow
        box = win.Toolkit.createMessageBox(win, boxtype, 1, title, message)
        return box.execute()

    def to_int(self, hex_str):
        if hex_str:
            return int(hex_str[-6:], 16)
        return 0

    def create_cfg_access(self):
        cfg = self.create('com.sun.star.configuration.ConfigurationProvider')
        prop = PropertyValue('nodepath', 0, '/ooo.ext.code-highlighter.Registry/Settings', 0)
        cfg_access = cfg.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess', (prop,))
        return cfg_access

    def load_options(self):
        properties = self.cfg_access.ElementNames
        values = self.cfg_access.getPropertyValues(properties)
        return dict(zip(properties, values))

    def save_options(self, **kwargs):
        self.options.update(kwargs)
        self.cfg_access.setPropertyValues(tuple(kwargs.keys()), tuple(kwargs.values()))
        self.cfg_access.commitChanges()

    def create_dialog(self):
        # get_all_lexers() returns:
        # (longname, tuple of aliases, tuple of filename patterns, tuple of mimetypes)
        all_lexers = sorted((lex[0] for lex in get_all_lexers()), key=str.casefold)
        self.all_lexer_aliases = [lex[0] for lex in get_all_lexers()]
        for lex in get_all_lexers():
            self.all_lexer_aliases.extend(list(lex[1]))
        self.all_styles = sorted(get_all_styles(), key=lambda x: (x != 'default', x.lower()))

        dialog_provider = self.create("com.sun.star.awt.DialogProvider")
        dialog = dialog_provider.createDialog("vnd.sun.star.extension://javahelps.codehighlighter/dialogs/CodeHighlighter.xdl")

        cb_lang = dialog.getControl('cb_lang')
        cb_style = dialog.getControl('cb_style')
        check_col_bg = dialog.getControl('check_col_bg')
        pygments_ver = dialog.getControl('pygments_ver')

        cb_lang.Text = self.options['Language']
        cb_lang.setSelection(Selection(0, len(cb_lang.Text)))
        cb_lang.addItems(all_lexers, 0)
        cb_lang.addItem('automatic', 0)

        style = self.options['Style']
        if style in self.all_styles:
            cb_style.Text = style
        cb_style.addItems(self.all_styles, 0)

        check_col_bg.State = int(self.options['ColorizeBackground'])

        def getextver():
            pip = self.ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
            extensions = pip.getExtensionList()
            for e in extensions:
                if "javahelps.codehighlighter" in e:
                    return e[1]
            return ''
        dialog.Title = dialog.Title.format(getextver())
        pygments_ver.Text = pygments_ver.Text.format(pygments.__version__)

        return dialog

    def getlexer(self, code):
        lang = self.options['Language']
        if lang == 'automatic':
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
        return lexer

    def highlight_source_code(self, selected_item=None):
        stylename = self.options['Style']
        style = styles.get_style_by_name(stylename)
        bg_color = style.background_color if self.options['ColorizeBackground'] != "0" else None

        self.doc.lockControllers()
        undomanager = self.doc.UndoManager
        hascode = False
        try:
            # Get the selected item
            if selected_item == None:
                selected_item = self.doc.CurrentSelection
            if not hasattr(selected_item, 'supportsService'):
                self.msgbox("Unsupported selection.")
                return
            elif hasattr(selected_item, 'getCount') and not hasattr(selected_item, 'queryContentCells'):
                for item_idx in range(selected_item.getCount()):
                    code_block = selected_item.getByIndex(item_idx)
                    code = code_block.String
                    if code.strip():
                        hascode = True
                        lexer = self.getlexer(code)
                        if code_block.supportsService('com.sun.star.drawing.Text'):
                            # TextBox
                            # exit edit mode if necessary
                            self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                            undoaction = UndoAction(self.doc, code_block, f"code highlight (lang: {lexer.name}, style: {stylename})")
                            code_block.FillStyle = FS_NONE
                            if bg_color:
                                code_block.FillStyle = FS_SOLID
                                code_block.FillColor = self.to_int(bg_color)
                            cursor = code_block.createTextCursorByRange(code_block)
                            cursor.CharLocale = self.nolocale
                            cursor.collapseToStart()
                            self.highlight_code(code, cursor, lexer, style)
                            # model is not considered as modified after textbox formatting
                            self.doc.setModified(True)
                            undoaction.get_new_state()
                            undomanager.addUndoAction(undoaction)
                        else:
                            # Plain text
                            try:
                                undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                                code_block.ParaBackColor = -1
                                if bg_color:
                                    code_block.ParaBackColor = self.to_int(bg_color)
                                cursor = code_block.getText().createTextCursorByRange(code_block)
                                cursor.CharLocale = self.nolocale
                                cursor.collapseToStart()
                                self.highlight_code(code, cursor, lexer, style)
                            finally:
                                undomanager.leaveUndoContext()

                if not hascode and selected_item.Count == 1:
                    code_block = selected_item[0]
                    if code_block.TextFrame:
                        self.highlight_source_code(code_block.TextFrame)
                    elif code_block.TextTable:
                        cellname = code_block.Cell.CellName
                        texttablecursor = code_block.TextTable.createCursorByCellName(cellname)
                        self.highlight_source_code(texttablecursor)
                    return


            elif selected_item.supportsService('com.sun.star.text.TextFrame'):
                # Selection is a text frame
                code_block = selected_item
                code = code_block.String
                if code.strip():
                    hascode = True
                    lexer = self.getlexer(code)
                    undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    try:
                        code_block.BackColor = -1
                        if bg_color:
                            code_block.BackColor = self.to_int(bg_color)
                        cursor = code_block.createTextCursorByRange(code_block)
                        cursor.CharLocale = self.nolocale
                        cursor.collapseToStart()
                        self.highlight_code(code, cursor, lexer, style)
                    finally:
                        undomanager.leaveUndoContext()

            elif selected_item.supportsService('com.sun.star.text.TextTableCursor'):
                # Selection is one or more table cell range
                table = self.doc.CurrentController.ViewCursor.TextTable
                rangename = selected_item.RangeName
                if ':' in rangename:
                    # at least two cells
                    cellrange = table.getCellRangeByName(rangename)
                    nrows, ncols = len(cellrange.Data), len(cellrange.Data[0])
                    for row in range(nrows):
                        for col in range(ncols):
                            code_block = cellrange.getCellByPosition(col, row)
                            code = code_block.String
                            if code.strip():
                                hascode = True
                                lexer = self.getlexer(code)
                                undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                                try:
                                    code_block.BackColor = -1
                                    if bg_color:
                                        code_block.BackColor = self.to_int(bg_color)
                                    cursor = code_block.createTextCursorByRange(code_block)
                                    cursor.CharLocale = self.nolocale
                                    cursor.collapseToStart()
                                    self.highlight_code(code, cursor, lexer, style)
                                finally:
                                    undomanager.leaveUndoContext()
                else:
                    # only one cell
                    code_block = table.getCellByName(rangename)
                    code = code_block.String
                    if code.strip():
                        hascode = True
                        lexer = self.getlexer(code)
                        undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                        try:
                            code_block.BackColor = -1
                            if bg_color:
                                code_block.BackColor = self.to_int(bg_color)
                            cursor = code_block.createTextCursorByRange(code_block)
                            cursor.CharLocale = self.nolocale
                            cursor.collapseToStart()
                            self.highlight_code(code, cursor, lexer, style)
                        finally:
                            undomanager.leaveUndoContext()

            elif selected_item.supportsService('com.sun.star.text.TextCursor'):
                # LO Impress shape selection
                cursor = selected_item
                code = cursor.String
                cdirection = cursor.compareRegionStarts(cursor.Start, cursor.End)
                if cdirection != 0:  # a selection exists
                    hascode = True
                    lexer = self.getlexer(code)
                    undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    try:
                        cursor.CharLocale = self.nolocale
                        if cdirection == 1:
                            cursor.collapseToStart()
                        else:
                            # if selection is done right to left inside text box, end cursor is before start cursor
                            cursor.collapseToEnd()

                        self.highlight_code(code, cursor, lexer, style)
                        # exit edit mode if necessary
                        self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                    finally:
                        undomanager.leaveUndoContext()

            elif hasattr(selected_item, 'queryContentCells'):
                # LO Calc cell selection
                self.dispatcher.executeDispatch(self.frame, ".uno:Deselect", "", 0, ())  # exit edit mode if necessary
                cells = selected_item.queryContentCells(CF_STRING).Cells
                if cells.hasElements():
                    hascode = True
                    for code_block in cells:
                        code = code_block.String
                        lexer = self.getlexer(code)
                        undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                        try:
                            code_block.CellBackColor = -1
                            code_block.CharLocale = self.nolocale
                            if bg_color:
                                code_block.CellBackColor = self.to_int(bg_color)
                            cursor = code_block.createTextCursor()
                            cursor.gotoStart(False)
                            self.highlight_code(code, cursor, lexer, style)
                        finally:
                            undomanager.leaveUndoContext()
            else:
                self.msgbox("Unsupported selection.")
                return

            if not hascode:
                self.msgbox("Nothing to highlight.")

        except Exception:
            self.msgbox(traceback.format_exc())
        finally:
            self.doc.unlockControllers()

    def highlight_code(self, code, cursor, lexer, style):
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
                    except Exception:
                        pass
                    finally:
                        cursor.collapseToEnd()  # deselects the selected text
                lastval = tok_value
                lasttype = tok_type


# Component registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(CodeHighlighter, "ooo.ext.code-highlighter.impl", (),)


# exposed functions for development stages
def highlight(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_highlight()


def highlight_previous(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_highlight_previous()
