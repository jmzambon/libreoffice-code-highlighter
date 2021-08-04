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

import uno
import unohelper
from com.sun.star.awt import Selection
from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD
from com.sun.star.awt.Key import RETURN as KEY_RETURN
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_YES_NO, DEFAULT_BUTTON_YES
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, ERRORBOX, INFOBOX, QUERYBOX
from com.sun.star.beans import PropertyValue
from com.sun.star.drawing.FillStyle import NONE as FS_NONE, SOLID as FS_SOLID
from com.sun.star.lang import Locale
from com.sun.star.sheet.CellFlags import STRING as CF_STRING
from com.sun.star.task import XJobExecutor


ctx = uno.getComponentContext()

def msgbox(message, title="Error", frame=None, boxtype=MESSAGEBOX, buttons=BUTTONS_OK):
    if not frame:
        desktop = ctx.ServiceManager.createInstance("com.sun.star.frame.Desktop")
        frame = desktop.ActiveFrame
    if frame.ActiveFrame:
        # top window is a subdocument
        frame = frame.ActiveFrame
    win = frame.ComponentWindow
    box = win.Toolkit.createMessageBox(win, boxtype, buttons, title, message)
    return box.execute()

def downloadpygments():
    import zipfile, urllib.request, ssl
    import json, shutil, importlib
    from sys import path as sys_path
    from os.path import join as os_path_join

    message = ("Pygments is not installed on this system, but is required\n"
               "for CodeHighlighter.\n\n"
               "Do you want to download it now?\n"
               "It will be installed within CodeHighlighter installation folder.")

    if msgbox(message, 
              boxtype=QUERYBOX,
              buttons=BUTTONS_YES_NO | DEFAULT_BUTTON_YES) == 3:   # YES = 2, NO = 3
        return False

    try:
        # prevent ssl errors
        requestcontext = ssl._create_unverified_context()

        # get last release url
        def getlastreleaseurl():
            baseurl = "https://api.github.com/repos/pygments/pygments/releases"
            with urllib.request.urlopen(baseurl, context=requestcontext) as response:
                if response.code == 200:
                    data = json.loads(response.read())
                    return data[0]['zipball_url']
                return None
        lastreleaseurl = getlastreleaseurl()

        # get extension absolute path
        def getextpath():
            pip = ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
            url = pip.getPackageLocation("javahelps.codehighlighter")
            return uno.fileUrlToSystemPath(url)
        extpath = getextpath()

        # grab latest release of pygments
        zipname = 'myzip'
        with urllib.request.urlopen(lastreleaseurl, context=requestcontext) as response, open(zipname, 'wb') as out_file:
            shutil.copyfileobj(response, out_file)
            with zipfile.ZipFile(zipname) as zf:
                infolist = zf.infolist()
                basename = infolist[0].filename 
                rep_pygments = [zipinfo for zipinfo in infolist[1:] if zipinfo.filename.startswith(basename + 'pygments/')]
                for file_ in rep_pygments:
                    file_.filename = file_.filename.replace(basename, 'python/pythonpath/', 1)
                    zf.extract(file_, extpath)

        # add pytonpath to syspath and import
        pygpath = os_path_join(extpath, 'python/pythonpath/')
        if not pygpath in sys_path:
            sys_path.append(pygpath)

        globals()['pygments'] = importlib.import_module('pygments')
        _styles = importlib.import_module('pygments.styles')
        globals()['styles'] = _styles
        globals()['get_all_styles'] = _styles.get_all_styles
        _lexers = importlib.import_module('pygments.lexers')
        globals()['get_all_lexers'] = _lexers.get_all_lexers
        globals()['get_lexer_by_name'] = _lexers.get_lexer_by_name
        globals()['guess_lexer'] = _lexers.guess_lexer

        return True

    except Exception:
        print('Error at pygments import:\n---')
        traceback.print_exc()
        msgbox("Sorry, unable to download pygments.", boxtype=ERRORBOX)


try:
    import pygments.util
    from pygments import styles
    from pygments.lexers import get_all_lexers
    from pygments.lexers import get_lexer_by_name
    from pygments.lexers import guess_lexer
    from pygments.styles import get_all_styles
except ImportError:
    # let CodeHighlighter take care of prompting for downloading pygments
    pass


class CodeHighlighter(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        print('CodeHighlighter.__init__()')
        try:
            self.ctx = ctx
            self.sm = ctx.ServiceManager
            self.desktop = self.create("com.sun.star.frame.Desktop")
            self.doc = self.desktop.getCurrentComponent()
            self.frame = self.desktop.ActiveFrame
            self.ensure_pygments = self.get_pygments_objects()
            if not self.ensure_pygments:
                return
            self.cfg_access = self.create_cfg_access()
            self.options = self.load_options()
            self.dialog = None
            self.dialog = self.create_dialog()
            self.nolocale = Locale("zxx", "", "")
        except Exception:
            traceback.print_exc()

    # XJobExecutor
    def trigger(self, arg):
        # print(f"trigger arg: {arg}")
        try:
            getattr(self, 'do_'+arg)()
        except:
            traceback.print_exc()

    # main functions
    def do_highlight(self):
        # get options choice        
        # 0: canceled, 1: OK
        # self.dialog.setVisible(True)
        if not self.ensure_pygments or self.dialog.execute() == 0:
            return
        lang = self.dialog.getControl('cb_lang').Text.strip() or 'automatic'
        style = self.dialog.getControl('cb_style').Text.strip() or 'default'
        colorize_bg = self.dialog.getControl('check_col_bg').State

        if lang != 'automatic' and lang not in self.all_lexer_aliases:
            msgbox("Unsupported language.", "Error", self.frame, ERRORBOX)
            return
        if style not in self.all_styles:
            msgbox("Unknown.", "Error", self.frame, ERRORBOX)
            return
        self.save_options(Style=style, Language=lang, ColorizeBackground=str(colorize_bg))
        self.highlight_source_code()

    def do_highlight_previous(self):
        if self.ensure_pygments:
            self.highlight_source_code()

    # private functions
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

    def load_options(self):
        properties = self.cfg_access.ElementNames
        values = self.cfg_access.getPropertyValues(properties)
        return dict(zip(properties, values))

    def save_options(self, **kwargs):
        self.options.update(kwargs)
        self.cfg_access.setPropertyValues(tuple(kwargs.keys()), tuple(kwargs.values()))
        self.cfg_access.commitChanges()

    def get_pygments_objects(self):
        try:
            # get_all_lexers() returns:
            # (longname, tuple of aliases, tuple of filename patterns, tuple of mimetypes)
            self.all_lexers = sorted((lex[0] for lex in get_all_lexers()), key=str.casefold)
            self.all_lexer_aliases = [lex[0] for lex in get_all_lexers()]
            for lex in get_all_lexers():
                self.all_lexer_aliases.extend(list(lex[1]))
            self.all_styles = sorted(get_all_styles(), key=lambda x: (x != 'default', x.lower()))
            return True
        except NameError:
            # if get_all_lexers does not exists -> pygments is not installed
            if not 'pygments' in globals():
                if downloadpygments() == True:
                    self.get_pygments_objects()
                    return True
                else:
                    return False
            else:
                raise

    def create_dialog(self):
        dialog_provider = self.create("com.sun.star.awt.DialogProvider")
        dialog = dialog_provider.createDialog("vnd.sun.star.extension://javahelps.codehighlighter/dialogs/CodeHighlighter.xdl")

        cb_lang = dialog.getControl('cb_lang')
        cb_style = dialog.getControl('cb_style')
        check_col_bg = dialog.getControl('check_col_bg')

        # TODO: reformat config access
        cb_lang.addItem('automatic', 0)
        cb_lang.Text = self.options['Language']
        cb_lang.setSelection(Selection(0, len(cb_lang.Text)))
        cb_lang.addItems(self.all_lexers, 0)

        style = self.options['Style']
        if style in self.all_styles:
            cb_style.Text = style
        cb_style.addItems(self.all_styles, 0)

        check_col_bg.State = int(self.options['ColorizeBackground'])

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

    def highlight_source_code(self):
        lang = self.options['Language']
        stylename = self.options['Style']
        style = styles.get_style_by_name(stylename)
        bg_color = style.background_color if self.options['ColorizeBackground'] != "0" else None

        self.doc.lockControllers()
        undomanager = self.doc.UndoManager
        try:
            # Get the selected item
            selected_item = self.doc.CurrentSelection
            if not hasattr(selected_item, 'supportsService'):
                return 
            elif hasattr(selected_item, 'getCount') and not hasattr(selected_item, 'queryContentCells'):
                for item_idx in range(selected_item.getCount()):
                    code_block = selected_item.getByIndex(item_idx)
                    code = code_block.String
                    if not code.strip():
                        continue
                    lexer = self.getlexer(code)
                    undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    try:
                        if code_block.supportsService('com.sun.star.drawing.Text'):
                            # TextBox
                            # exit edit mode if necessary
                            dispatcher = self.create("com.sun.star.frame.DispatchHelper")
                            dispatcher.executeDispatch(self.doc.CurrentController.Frame, ".uno:SelectObject", "", 0, ())
                            code_block.FillStyle = FS_NONE
                            if bg_color:
                                code_block.FillStyle = FS_SOLID
                                code_block.FillColor = self.to_int(bg_color)
                            cursor = code_block.createTextCursorByRange(code_block)
                            cursor.CharLocale = self.nolocale
                            cursor.collapseToStart()
                        else:
                            # Plain text
                            code_block.ParaBackColor = -1
                            if bg_color:
                                code_block.ParaBackColor = self.to_int(bg_color)
                            cursor = code_block.getText().createTextCursorByRange(code_block)
                            cursor.CharLocale = self.nolocale
                            cursor.collapseToStart()
                        self.highlight_code(code, cursor, lexer, style)
                    finally:
                        undomanager.leaveUndoContext()

            elif selected_item.supportsService('com.sun.star.text.TextFrame'):
                # Selection is a text frame
                code_block = selected_item
                code = code_block.String
                if not code.strip():
                    return
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
                            if not code.strip():
                                continue
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
                    if not code.strip():
                        return
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
                if cdirection == 0:  # no selection
                    return
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
                    dispatcher = self.create("com.sun.star.frame.DispatchHelper")
                    dispatcher.executeDispatch(self.doc.CurrentController.Frame, ".uno:SelectObject", "", 0, ())
                finally:
                    undomanager.leaveUndoContext()

            elif hasattr(selected_item, 'queryContentCells'):
                # LO Calc cell selection
                cells = selected_item.queryContentCells(CF_STRING).Cells
                if cells.hasElements():
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
        except Exception:
            msgbox(traceback.format_exc(), frame=self.frame)
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
                    except:
                        pass
                    finally:
                        cursor.collapseToEnd()  # deselects the selected text
                lastval = tok_value
                lasttype = tok_type


# Component registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(CodeHighlighter, "ooo.ext.code-highlighter.impl", (),)


#--------------------------------------------------------------------------------
# exposed functions for development stages
#--------------------------------------------------------------------------------
def highlight(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx) 
    highlighter.do_highlight()

def highlight_previous(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx) 
    highlighter.do_highlight_previous()


g_exportedScripts = highlight, highlight_previous
