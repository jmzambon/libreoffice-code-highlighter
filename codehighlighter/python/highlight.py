# Code Highligher 2 is a LibreOffice extension to highlight code snippets
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

# python standard
import re
import traceback
from math import log10

# uno
import unohelper
from com.sun.star.awt import Selection
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, ERRORBOX
from com.sun.star.beans import PropertyValue
from com.sun.star.drawing.FillStyle import NONE as FS_NONE, SOLID as FS_SOLID
from com.sun.star.lang import Locale
from com.sun.star.sheet.CellFlags import STRING as CF_STRING
from com.sun.star.task import XJobExecutor
from com.sun.star.document import XUndoAction

# python standard
import pygments
from pygments import styles
from pygments.lexers import get_all_lexers
from pygments.lexers import get_lexer_by_name
from pygments.lexers import guess_lexer
from pygments.styles import get_all_styles

# internal
import ch2_i18n


class UndoAction(unohelper.Base, XUndoAction):
    """ Add undo/redo action for highlighting not catched by the system,
        i.e. when applied on textbox objects."""

    def __init__(self, doc, textbox, title):
        self.doc = doc
        self.textbox = textbox
        self.Title = title
        self.old_portions = None
        self.old_bg = None
        self.new_portions = None
        self.new_bg = None
        self.charprops = ("CharColor", "CharLocale", "CharPosture", "CharHeight", "CharWeight")
        self.bgprops = ("FillColor", "FillStyle")
        self.get_old_state()

    def undo(self):
        self.textbox.setString(self.old_text)
        self._format(self.old_portions, self.old_bg)

    def redo(self):
        self.textbox.setString(self.new_text)
        self._format(self.new_portions, self.new_bg)

    def get_old_state(self):
        self.old_bg = self.textbox.getPropertyValues(self.bgprops)
        self.old_text = self.textbox.String
        self.old_portions = self._extract_portions()

    def get_new_state(self):
        self.new_bg = self.textbox.getPropertyValues(self.bgprops)
        self.new_text = self.textbox.String
        self.new_portions = self._extract_portions()

    def _extract_portions(self):
        textportions = []
        for para in self.textbox:
            if textportions:    # new paragraph after first one
                textportions[-1][0] += 1
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


class CodeHighlighter(unohelper.Base, XJobExecutor, XDialogEventHandler):
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
            self.strings = ch2_i18n.getstrings(ctx)
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

    # XDialogEventHandler
    def callHandlerMethod(self, dialog, event, method):
        if method == "topage1":
            dialog.Model.Step = 1
            dialog.getControl('cb_lang').setFocus()
            return True
        elif method == "topage2":
            dialog.Model.Step = 2
            dialog.getControl('nb_start').setFocus()
            return True
        return False

    def getSupportedMethodNames(self):
        return 'topage1', 'topage2'

    # main functions
    def do_highlight(self):
        if self.choose_options():
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

    def create_dialog(self):
        # get_all_lexers() returns:
        # (longname, tuple of aliases, tuple of filename patterns, tuple of mimetypes)
        all_lexers = sorted((lex[0] for lex in get_all_lexers()), key=str.casefold)
        self.all_lexer_aliases = [lex[0].lower() for lex in get_all_lexers()]
        for lex in get_all_lexers():
            self.all_lexer_aliases.extend(list(lex[1]))
        self.all_styles = sorted(get_all_styles(), key=lambda x: (x != 'default', x.lower()))

        dialog_provider = self.create("com.sun.star.awt.DialogProvider2")
        dialog = dialog_provider.createDialogWithHandler("vnd.sun.star.extension://javahelps.codehighlighter/dialogs/CodeHighlighter2.xdl", self)

        # set localized strings
        controlnames = ("label_lang", "label_style", "check_col_bg", "check_linenb", "nb_line", "lbl_nb_start",
                        "lbl_nb_ratio", "lbl_nb_sep", "lbl_nb_spaces", "pygments_ver", "topage1", "topage2")
        for controlname in controlnames:
            dialog.getControl(controlname).Model.setPropertyValues(("Label", "HelpText"), self.strings[controlname])
        for controlname in ("nb_sep", "nb_spaces"):
            dialog.getControl(controlname).Model.HelpText = self.strings[controlname][1]

        cb_lang = dialog.getControl('cb_lang')
        cb_style = dialog.getControl('cb_style')
        check_col_bg = dialog.getControl('check_col_bg')
        check_linenb = dialog.getControl('check_linenb')
        nb_start = dialog.getControl('nb_start')
        nb_ratio = dialog.getControl('nb_ratio')
        nb_sep = dialog.getControl('nb_sep')
        nb_spaces = dialog.getControl('nb_spaces')
        pygments_ver = dialog.getControl('pygments_ver')

        cb_lang.Text = self.options['Language']
        cb_lang.setSelection(Selection(0, len(cb_lang.Text)))
        cb_lang.addItems(all_lexers, 0)
        cb_lang.addItem('automatic', 0)

        style = self.options['Style']
        if style in self.all_styles:
            cb_style.Text = style
        cb_style.addItems(self.all_styles, 0)

        check_col_bg.State = self.options['ColorizeBackground']
        check_linenb.State = self.options['ShowLineNumbers']
        nb_start.Value = self.options['LineNumberStart']
        nb_ratio.Value = self.options['LineNumberRatio']
        nb_sep.Text = self.options['LineNumberSeparator']
        nb_spaces.Text = self.options['LineNumberSpaces']

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

    def load_options(self):
        properties = self.cfg_access.ElementNames
        values = self.cfg_access.getPropertyValues(properties)
        return dict(zip(properties, values))

    def choose_options(self):
        # get options choice
        # 0: canceled, 1: OK
        # self.dialog.setVisible(True)
        if self.dialog.execute() == 0:
            return False
        lang = self.dialog.getControl('cb_lang').Text.strip() or 'automatic'
        style = self.dialog.getControl('cb_style').Text.strip() or 'default'
        colorize_bg = self.dialog.getControl('check_col_bg').State
        show_linenb = self.dialog.getControl('check_linenb').State
        nb_start = int(self.dialog.getControl('nb_start').Value)
        nb_ratio = int(self.dialog.getControl('nb_ratio').Value)
        nb_sep = self.dialog.getControl('nb_sep').Text
        nb_spaces = self.dialog.getControl('nb_spaces').Text


        if lang != 'automatic' and lang.lower() not in self.all_lexer_aliases:
            self.msgbox(self.strings["errlang"])
            return False
        if style not in self.all_styles:
            self.msgbox(self.strings["errstyle"])
            return False
        self.save_options(Style=style, Language=lang, ColorizeBackground=colorize_bg, ShowLineNumbers=show_linenb,
                          LineNumberStart=nb_start, LineNumberRatio=nb_ratio, LineNumberSeparator=nb_sep, LineNumberSpaces=nb_spaces)
        return True     

    def save_options(self, **kwargs):
        self.options.update(kwargs)
        self.cfg_access.setPropertyValues(tuple(kwargs.keys()), tuple(kwargs.values()))
        self.cfg_access.commitChanges()

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
                    if lex[0].lower() == lang.lower():
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
        bg_color = style.background_color if self.options['ColorizeBackground'] else None

        self.doc.lockControllers()
        undomanager = self.doc.UndoManager
        hascode = False
        try:
            # Get the selected item
            if selected_item == None:
                selected_item = self.doc.CurrentSelection
            if not hasattr(selected_item, 'supportsService'):
                self.msgbox(self.strings["errsel1"])
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
                            if self.show_line_numbers(code_block):
                                code = code_block.String    #code string has changed
                            cursor = code_block.createTextCursorByRange(code_block)
                            cursor.CharLocale = self.nolocale
                            cursor.collapseToStart()
                            self.highlight_code(code, cursor, lexer, style)
                            # unlock controllers here to force left pane syncing in draw/impress
                            if self.doc.supportsService("com.sun.star.drawing.GenericDrawingDocument"):
                                self.doc.unlockControllers()
                            code_block.FillStyle = FS_NONE
                            if bg_color:
                                code_block.FillStyle = FS_SOLID
                                code_block.FillColor = self.to_int(bg_color)
                            # model is not considered as modified after textbox formatting
                            self.doc.setModified(True)
                            undoaction.get_new_state()
                            undomanager.addUndoAction(undoaction)
                        else:
                            # Plain text
                            try:
                                undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                                self.show_line_numbers(code_block, isplaintext=True)
                                cursor, code = self.ensure_selection(code_block)
                                cursor.ParaBackColor = -1
                                if bg_color:
                                    cursor.ParaBackColor = self.to_int(bg_color)
                                cursor.CharLocale = self.nolocale
                                self.doc.CurrentController.select(cursor)
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
                    if self.show_line_numbers(code_block):
                        code = code_block.String    #code string has changed
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
                                if self.show_line_numbers(code_block):
                                    code = code_block.String    #code string has changed
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
                        if self.show_line_numbers(code_block):
                            code = code_block.String    #code string has changed
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
                # LO Impress text selection inside shape -> highlight all shape
                # exit edit mode
                self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                self.highlight_source_code()
                return

                ### OLD CODE, intended to highlight sub text, but api's too buggy'
                # # first exit edit mode, otherwise formatting is not shown (bug?)
                # self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                # cursor = selected_item
                # code = cursor.String
                # cdirection = cursor.compareRegionStarts(cursor.Start, cursor.End)
                # if cdirection != 0:  # a selection exists
                #     hascode = True
                #     lexer = self.getlexer(code)
                #     undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                #     try:
                #         cursor.CharBackColor = -1
                #         if bg_color:
                #             cursor.CharBackColor = self.to_int(bg_color)
                #         cursor.CharLocale = self.nolocale
                #         if cdirection == 1:
                #             cursor.collapseToStart()
                #         else:
                #             # if selection is done right to left inside text box, end cursor is before start cursor
                #             cursor.collapseToEnd()
                #         self.highlight_code(code, cursor, lexer, style)
                #     finally:
                #         undomanager.leaveUndoContext()

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
                        if self.show_line_numbers(code_block):
                            code = code_block.String    #code string has changed
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
                self.msgbox(self.strings["errsel1"])
                return

            if not hascode:
                self.msgbox(self.strings["errsel2"])

        except Exception:
            self.msgbox(traceback.format_exc())
        finally:
            if self.doc.hasControllersLocked():
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

    def show_line_numbers(self, code_block, isplaintext=False):
        show_linenb = self.options['ShowLineNumbers']
        startnb = self.options["LineNumberStart"]
        ratio = self.options["LineNumberRatio"]
        spaces = self.options["LineNumberSpaces"]
        if spaces == r'\t':
            spaces = '\t'
        sep = self.options["LineNumberSeparator"]  # allowed values: ".", ":" or ""
        if not sep in (".:"):
            sep = ""
        charsize = code_block.End.CharHeight
        numbersize = round(charsize*ratio//50)/2   # round to 0.5

        p = re.compile(f"\s*([0-9]+)[\.|:]?{spaces}")
        c = code_block.Text.createTextCursor()
        code = c.Text.String
        if isplaintext:
            # if cursor is not at start of paragraph for plain text
            # selection, line numbering can't be detected. So let's
            # move the cursor to the whole first paragraph.
            c = code_block.Text.createTextCursorByRange(code_block)
            c.gotoStartOfParagraph(False)
            c.gotoEndOfParagraph(True)
            code = c.String

        def show_numbering():
            nblignes = len(code_block.String.split('\n'))
            digits = int(log10(nblignes - 1 + startnb)) + 1
            for n, para in enumerate(code_block, start=startnb):
                # para.Start.CharHeight = numbersize
                prefix = f'{n:>{digits}}{sep}{spaces}'
                para.Start.setString(prefix)
                c.gotoRange(para.Start, False)
                c.goRight(len(prefix), True)
                c.CharHeight = numbersize

        def hide_numbering():
            for para in code_block:
                m = p.match(para.String)
                if m:
                    c.gotoRange(para.Start, False)
                    c.goRight(m.end(), True)
                    c.CharHeight = charsize
                    c.setString("")

        m = p.match(code) 
        res = False
        if show_linenb:
            if not m:
                show_numbering()
                res = True
            else:
                # numbering already exists, but let's replace it anyway,
                # as new settings can have been defined.
                hide_numbering()
                show_numbering()
                res = True
        elif m:
            hide_numbering()
            res = True
        return res

    def ensure_selection(self, selected_code):
        c = selected_code.Text.createTextCursorByRange(selected_code)
        c.gotoStartOfParagraph(False)
        c.gotoRange(selected_code.End, True)
        c.gotoEndOfParagraph(True)
        return c, c.String


# Component registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(CodeHighlighter, "ooo.ext.code-highlighter.impl", (),)


# exposed functions for development stages only
# uncomment corresponding entry in manifest.xml to add them as framework scripts
def highlight(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_highlight()


def highlight_previous(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_highlight_previous()
