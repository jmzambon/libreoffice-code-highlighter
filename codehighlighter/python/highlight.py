# Code Highligher 2 is a LibreOffice extension to highlight code snippets
# over 350 languages.

# Copyright (C) 2017  Gobinath, 2022 jmzambon

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

# prepare logger
import uno
import os.path
import logging
from com.sun.star.uno import RuntimeException
LOGLEVEL = {0: logging.WARNING, 1: logging.INFO, 2: logging.DEBUG}
logger = logging.getLogger("codehighlighter")
formatter = logging.Formatter("%(levelname)s [%(funcName)s::%(lineno)d] %(message)s")
consolehandler = logging.StreamHandler()
consolehandler.setFormatter(formatter)
try:
    userpath = uno.getComponentContext().ServiceManager.createInstance(
                    "com.sun.star.util.PathSubstitution").substituteVariables("$(user)", True)
    logfile = os.path.join(uno.fileUrlToSystemPath(userpath), "codehighlighter.log")
    filehandler = logging.FileHandler(logfile, mode="w", delay=True)
    filehandler.setFormatter(formatter)
    logger.addHandler(consolehandler)
    logger.setLevel(logging.INFO)
    logger.info("Logger installed.")
except RuntimeException:
    # At installation time, no context is available -> just ignore it.
    pass

# simple import hook, making sure embedded pygments is found first
import sys
try:
    path = os.path.join(os.path.dirname(__file__), "pythonpath")
    sys.path.insert(0, sys.path.pop(sys.path.index(path)))
    logger.debug(f'sys.path: {sys.path}')
    logger.info("Embedded Pygments path priorised.")
except NameError:
    # __file__ is not defined
    # only occurs when using exposed functions -> should be harmless
    pass
except Exception as e:
    logger.exception("")

try:
    # python standard
    import re
    import traceback
    from math import log10
    from ast import literal_eval

    # pygments
    import pygments
    from pygments.lexers import get_all_lexers, get_lexer_by_name, guess_lexer
    from pygments.styles import get_all_styles, get_style_by_name
    logger.info(f"Pygments imported from {pygments.__file__}.")
    logger.info(f"Lexers imported from {pygments.lexers.__file__}.")

    # uno
    import unohelper
    from com.sun.star.awt import Selection, XDialogEventHandler
    from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD
    from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
    from com.sun.star.awt.FontUnderline import NONE as UL_NONE, SINGLE as UL_SINGLE
    from com.sun.star.awt.MessageBoxType import ERRORBOX
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import ElementExistException
    from com.sun.star.document import XUndoAction
    from com.sun.star.drawing.FillStyle import SOLID as FS_SOLID  # , NONE as FS_NONE
    from com.sun.star.lang import Locale
    from com.sun.star.sheet.CellFlags import STRING as CF_STRING
    from com.sun.star.task import XJobExecutor
    from com.sun.star.xml import AttributeData

    # internal
    import ch2_i18n
except Exception:
    logger.exception("")


CHARSTYLEID = "ch2_"
SNIPPETTAGID = CHARSTYLEID + "options"


class UndoAction(unohelper.Base, XUndoAction):
    '''
    Add undo/redo action for highlighting operations not catched by the system,
    i.e. when applied on textbox objects.
    '''

    def __init__(self, doc, textbox, title):
        self.doc = doc
        self.textbox = textbox
        self.old_portions = None
        self.old_bg = None
        self.new_portions = None
        self.new_bg = None
        self.charprops = ("CharBackColor", "CharColor", "CharLocale", "CharPosture", "CharHeight", "CharWeight")
        self.bgprops = ("FillColor", "FillStyle")
        self.get_old_state()
        # XUndoAction attribute
        self.Title = title

    # XUndoAction (https://www.openoffice.org/api/docs/common/ref/com/sun/star/document/XUndoAction.html)
    def undo(self):
        self.textbox.setString(self.old_text)
        self.textbox.UserDefinedAttributes = self.old_attributes
        self._format(self.old_portions, self.old_bg)

    def redo(self):
        self.textbox.setString(self.new_text)
        self.textbox.UserDefinedAttributes = self.new_attributes
        self._format(self.new_portions, self.new_bg)

    # public
    def get_old_state(self):
        '''
        Gather text formattings before code highlighting.
        Will be used by <undo> to restore old state.
        '''

        self.old_bg = self.textbox.getPropertyValues(self.bgprops)
        self.old_text = self.textbox.String
        self.old_portions = self._extract_portions()
        self.old_attributes = self.textbox.UserDefinedAttributes

    def get_new_state(self):
        '''
        Gather text formattings after code highlighting.
        Will be used by <redo> to apply new state again.
        '''

        self.new_bg = self.textbox.getPropertyValues(self.bgprops)
        self.new_text = self.textbox.String
        self.new_portions = self._extract_portions()
        self.new_attributes = self.textbox.UserDefinedAttributes

    # private
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
            self.charstylesavailable = (
                    'CharacterStyles' in self.doc.StyleFamilies and
                    self.doc.CurrentSelection.ImplementationName != "com.sun.star.drawing.SvxShapeCollection")
            self.cfg_access = self.create_cfg_access()
            self.options = self.load_options()
            self.setlogger()
            logger.debug(f"Code Highlighter started from {self.doc.Title}.")
            logger.info(f"Using Pygments version {pygments.__version__}.")
            logger.info(f"Loaded options = {self.options}.")
            self.frame = self.doc.CurrentController.Frame
            self.dispatcher = self.create("com.sun.star.frame.DispatchHelper")
            self.strings = ch2_i18n.getstrings(ctx)
            self.nolocale = Locale("zxx", "", "")
            self.inlinesnippet = False
        except Exception:
            logger.exception("")

    # XJobExecutor (https://www.openoffice.org/api/docs/common/ref/com/sun/star/task/XJobExecutor.html)
    def trigger(self, arg):
        logger.debug(f"Code Highlighter triggered with argument '{arg}'.")
        try:
            getattr(self, 'do_'+arg)()
        except Exception:
            logger.exception("")

    # XDialogEventHandler (http://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/XDialogEventHandler.html)
    def callHandlerMethod(self, dialog, event, method):
        logger.debug(f"Dialog handler action: '{method}'.")
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
        '''Open option dialog and start code highlighting.'''

        if self.choose_options():
            self.prepare_highlight()

    def do_highlight_previous(self):
        '''Start code highlighting with current options as default.'''

        self.prepare_highlight()

    def do_update(self):
        '''Update already highlighted snippets based on options stored in codeblock tags.
        Code-blocks must have been highlighted at least once with Code Highlighter 2.'''

        self.prepare_highlight(updatecode=True)

    def do_removealltags(self):
        '''Remove all highlighting infos inserted with Code Highlighter 2
        in the active document.'''

        self.removealltags()

    # private functions
    def create(self, service):
        '''Instanciate UNO services'''

        return self.sm.createInstance(service)

    def msgbox(self, message, boxtype=ERRORBOX, title="Error"):
        '''Simple UNO message box for notifications at user.'''

        win = self.frame.ContainerWindow
        box = win.Toolkit.createMessageBox(win, boxtype, 1, title, message)
        return box.execute()

    def to_int(self, hex_str):
        '''Convert hexadecimal color representation into decimal integer.'''

        if hex_str:
            return int(hex_str[-6:], 16)
        return 0

    def setlogger(self):
        loglevel = LOGLEVEL.get(self.options["LogLevel"], 0)
        logger.setLevel(loglevel)
        if self.options["LogToFile"] == 0:
            logger.removeHandler(filehandler)
            logger.addHandler(consolehandler)
        else:
            logger.removeHandler(consolehandler)
            logger.addHandler(filehandler)

    def create_cfg_access(self):
        '''Return an updatable instance of the codehighlighter node in LO registry. '''

        cfg = self.create('com.sun.star.configuration.ConfigurationProvider')
        prop = PropertyValue('nodepath', 0, '/ooo.ext.code-highlighter.Registry/Settings', 0)
        cfg_access = cfg.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess', (prop,))
        return cfg_access

    def load_options(self):
        properties = self.cfg_access.ElementNames
        values = self.cfg_access.getPropertyValues(properties)
        return dict(zip(properties, values))

    def getallstyles(self):
        all_styles = list(get_all_styles()) + ['libreoffice-classic', 'libreoffice-dark']
        return sorted(all_styles, key=lambda x: (x != 'default', x.casefold()))

    def create_dialog(self):
        '''Load, populate and return options dialog.'''

        # get_all_lexers() returns: (longname, tuple of aliases, tuple of filename patterns, tuple of mimetypes)
        logger.debug("Starting options dialog.")
        _all_lexers = list(get_all_lexers())
        # let's add a convenient shortcut to VB.net lexer for LOBasic
        _all_lexers.append(("LibreOffice Basic", (), (), ()))
        all_lexers = sorted((lex[0] for lex in _all_lexers), key=str.casefold)
        self.all_lexer_aliases = [lex[0].lower() for lex in _all_lexers]
        for lex in _all_lexers:
            self.all_lexer_aliases.extend(list(lex[1]))
        logger.debug("--> getting lexers ok.")
        self.all_styles = self.getallstyles()
        logger.debug("--> getting styles ok.")

        dialog_provider = self.create("com.sun.star.awt.DialogProvider2")
        dialog = dialog_provider.createDialogWithHandler(
            "vnd.sun.star.extension://javahelps.codehighlighter/dialogs/CodeHighlighter2.xdl", self)
        logger.debug("--> creating dialog ok.")

        # set localized strings
        controlnames = ("label_lang", "label_style", "check_col_bg", "check_charstyles", "check_linenb",
                        "nb_line", "cs_line", "lbl_nb_start", "lbl_nb_ratio", "lbl_nb_sep", "lbl_nb_pad",
                        "lbl_cs_rootstyle", "pygments_ver", "topage1", "topage2")
        for controlname in controlnames:
            dialog.getControl(controlname).Model.setPropertyValues(("Label", "HelpText"), self.strings[controlname])
        controlnames = ("nb_sep", "nb_pad", "cs_rootstyle")
        for controlname in controlnames:
            dialog.getControl(controlname).Model.HelpText = self.strings["lbl_" + controlname][1]

        cb_lang = dialog.getControl('cb_lang')
        cb_style = dialog.getControl('cb_style')
        check_col_bg = dialog.getControl('check_col_bg')
        check_linenb = dialog.getControl('check_linenb')
        check_charstyles = dialog.getControl('check_charstyles')
        check_charstyles.setEnable(self.charstylesavailable)
        nb_start = dialog.getControl('nb_start')
        nb_ratio = dialog.getControl('nb_ratio')
        nb_sep = dialog.getControl('nb_sep')
        nb_pad = dialog.getControl('nb_pad')
        cs_rootstyle = dialog.getControl('cs_rootstyle')
        pygments_ver = dialog.getControl('pygments_ver')

        cb_lang.Text = self.options['Language']
        cb_lang.setSelection(Selection(0, len(cb_lang.Text)))
        cb_lang.addItems(all_lexers, 0)
        cb_lang.addItem('automatic', 0)

        style = self.options['Style']
        if style in self.all_styles:
            cb_style.Text = style
        cb_style.addItems(self.all_styles, 0)

        check_col_bg.State = self.options['ColourizeBackground']
        check_charstyles.State = self.options['UseCharStyles']
        check_linenb.State = self.options['ShowLineNumbers']
        nb_start.Value = self.options['LineNumberStart']
        nb_ratio.Value = self.options['LineNumberRatio']
        nb_sep.Text = self.options['LineNumberSeparator']
        nb_pad.Text = self.options['LineNumberPaddingSymbol']
        cs_rootstyle.Text = self.options['MasterCharStyle']
        logger.debug("--> filling controls ok.")

        def getextver():
            pip = self.ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
            extensions = pip.getExtensionList()
            for e in extensions:
                if "javahelps.codehighlighter" in e:
                    return e[1]
            return ''
        dialog.Title = dialog.Title.format(getextver())
        pygments_ver.Text = pygments_ver.Text.format(pygments.__version__)
        logger.debug("Dialog returned.")

        return dialog

    def choose_options(self):
        '''
        Get options choice.
        Dialog return values: 0 = Canceled, 1 = OK
        '''

        # dialog.setVisible(True)
        dialog = self.create_dialog()
        if dialog.execute() == 0:
            logger.debug("Dialog canceled.")
            return False
        lang = dialog.getControl('cb_lang').Text.strip() or 'automatic'
        style = dialog.getControl('cb_style').Text.strip() or 'default'
        colorize_bg = dialog.getControl('check_col_bg').State
        use_charstyles = dialog.getControl('check_charstyles').State
        show_linenb = dialog.getControl('check_linenb').State
        nb_start = int(dialog.getControl('nb_start').Value)
        nb_ratio = int(dialog.getControl('nb_ratio').Value)
        nb_sep = dialog.getControl('nb_sep').Text
        nb_pad = dialog.getControl('nb_pad').Text
        cs_rootstyle = dialog.getControl('cs_rootstyle').Text

        if lang != 'automatic' and lang.lower() not in self.all_lexer_aliases:
            self.msgbox(self.strings["errlang"])
            return False
        if style not in self.all_styles:
            self.msgbox(self.strings["errstyle"])
            return False
        self.save_options(Style=style, Language=lang, ColourizeBackground=colorize_bg, ShowLineNumbers=show_linenb,
                          LineNumberStart=nb_start, LineNumberRatio=nb_ratio, LineNumberSeparator=nb_sep,
                          LineNumberPaddingSymbol=nb_pad, UseCharStyles=use_charstyles, MasterCharStyle=cs_rootstyle)
        logger.debug("Dialog validated and options saved.")
        logger.info(f"Updated options = {self.options}.")
        return True

    def save_options(self, **kwargs):
        self.options.update(kwargs)
        self.cfg_access.setPropertyValues(tuple(kwargs.keys()), tuple(kwargs.values()))
        self.cfg_access.commitChanges()

    def getlexerbyname(self, lexername):
        if lexername == 'LibreOffice Basic':
            lexername = "VB.net"
        try:
            return get_lexer_by_name(lexername)
        except pygments.util.ClassNotFound:
            # get_lexer_by_name() only checks aliases, not the actual longname
            for lex in get_all_lexers():
                if lex[0].lower() == lexername.lower():
                    # found the longname, use the first alias
                    return get_lexer_by_name(lex[1][0])
            else:
                raise

    def guesslexer(self, code_block):
        try:
            udas = code_block.UserDefinedAttributes
        except AttributeError:
            if self.inlinesnippet:
                udas = code_block.TextUserDefinedAttributes
            else:
                udas = code_block.ParaUserDefinedAttributes
        except Exception:
            logger.exception("")
            return guess_lexer(code_block.String)
        if udas is None or SNIPPETTAGID not in udas:
            return guess_lexer(code_block.String)
        else:
            options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
            logger.info('lexer name gotten from from snippet tag')
            if options['Language'] == "Text only":
                return guess_lexer(code_block.String)
            else:
                return self.getlexerbyname(options['Language'])

    def getlexer(self, code_block):
        lang = self.options['Language']
        if lang == 'automatic':
            lexer = self.guesslexer(code_block)
            logger.info(f'Automatic lexer choice : {lexer.name}')
        else:
            lexer = self.getlexerbyname(lang)
        # prevent offset color if selection start with empty line
        lexer.stripnl = False
        return lexer

    def createcharstyles(self, style, styleprefix):
        def addstyle(ttype):
            newcharstyle = self.doc.createInstance("com.sun.star.style.CharacterStyle")
            ttypename = str(ttype).replace('Token', styleprefix)
            try:
                charstyles.insertByName(ttypename, newcharstyle)
            except ElementExistException:
                return
            if ttype.parent is not None:
                parent = ttypename.rsplit('.', 1)[0]
                if not charstyles.hasByName(parent):
                    addstyle(ttype.parent)
                newcharstyle.ParentStyle = parent
            elif mastercharstyle:
                if not charstyles.hasByName(mastercharstyle):
                    master = self.doc.createInstance("com.sun.star.style.CharacterStyle")
                    charstyles.insertByName(mastercharstyle, master)
                newcharstyle.ParentStyle = mastercharstyle
            for d in style.styles.get(ttype, '').split():
                tok_style = style.style_for_token(ttype)
                if d == "noinherit":
                    break
                elif d == 'italic':
                    newcharstyle.CharPosture = SL_ITALIC
                elif d == 'noitalic':
                    newcharstyle.CharPosture = SL_NONE
                elif d == 'bold':
                    newcharstyle.CharWeight = W_BOLD
                elif d == 'nobold':
                    newcharstyle.CharWeight = W_NORMAL
                elif d == 'underline':
                    newcharstyle.CharUnderline = UL_SINGLE
                elif d == 'nounderline':
                    newcharstyle.CharUnderline = UL_NONE
                elif d in ('roman', 'sans', 'mono'):
                    pass
                elif d.startswith('bg:'):
                    # let's Pygments make the hard job here
                    if tok_style["bgcolor"]:
                        newcharstyle.CharBackColor = self.to_int(tok_style["bgcolor"])
                elif d.startswith('border:'):
                    pass
                elif d:
                    # let's Pygments make the hard job here
                    if tok_style["color"]:
                        newcharstyle.CharColor = self.to_int(tok_style["color"])

        mastercharstyle = self.options["MasterCharStyle"].strip()
        stylefamilies = self.doc.StyleFamilies
        charstyles = stylefamilies.CharacterStyles
        for ttype in sorted(style.styles.keys()):
            addstyle(ttype)

    def cleancharstyles(self, styleprefix):
        try:
            stylefamilies = self.doc.StyleFamilies
            charstyles = stylefamilies.CharacterStyles
            for cs in charstyles.ElementNames:
                # Remove only the styles created with certainty by the extension
                if cs.startswith(CHARSTYLEID) or cs.startswith(f'{styleprefix}.'):
                    if not charstyles.getByName(cs).isInUse():
                        charstyles.removeByName(cs)
        except AttributeError:
            pass

    def getstylebyname(self, name):
        if name.startswith('libreoffice'):
            from customstyles import libreoffice
            libostyles = {'libreoffice-classic': 'LibreOfficeStyle', 'libreoffice-dark': 'LibreOfficeDarkStyle'}
            return getattr(libreoffice, libostyles[name])
        else:
            return get_style_by_name(name)

    def tagcodeblock(self, code_block, lexername):
        if not self.options["StoreOptionsWithSnippet"]:
            return
        if self.inlinesnippet:
            logger.info('Code identified as inline snippet.')
            udas = code_block.TextUserDefinedAttributes
        else:
            try:
                udas = code_block.UserDefinedAttributes
            except AttributeError:
                udas = code_block.ParaUserDefinedAttributes
            except Exception:
                logger.exception("")
                return
        if udas is not None:
            _options = {k: self.options[k] for k in self.options if not k.startswith('Log')}
            _options['Language'] = lexername
            options = AttributeData(Type="CDATA", Value=f'{_options}')
            try:
                udas.insertByName(SNIPPETTAGID, options)
            except ElementExistException:
                udas.replaceByName(SNIPPETTAGID, options)
            if self.inlinesnippet:
                code_block.TextUserDefinedAttributes = udas
            else:
                try:
                    code_block.UserDefinedAttributes = udas
                    logger.info(f'snippet tagged with options: {_options}')
                except AttributeError:
                    code_block.ParaUserDefinedAttributes = udas
        else:
            logger.debug("Problem while saving user defined attributes: code block is probably mixing attributes. ")
            logger.debug(f"Code block concerned: {code_block.String}")

    def prepare_highlight(self, selected_item=None, updatecode=False):
        '''
        Check if selection is valid and contains text.
        If there is no selection but cursor is inside a text frame or
        a text table cell, and that this frame or cell contains text,
        selection is extended to the whole container.
        If cursor is inside a text shape or a Calc cell, selection is extended
        to the whole container in any case.
        If selection contains only part of paragraphs, selection is
        extended to the entire paragraphs.
        '''

        stylename = self.options['Style']
        style = self.getstylebyname(stylename)
        bg_color = style.background_color if self.options['ColourizeBackground'] else None
        if self.options['ShowLineNumbers']:
            _style = style.style_for_token(("Comment",))
            lineno_color = self.to_int(_style['color'])
        if not self.doc.hasControllersLocked():
            self.doc.lockControllers()
            logger.debug("Controllers locked.")
        undomanager = self.doc.UndoManager
        hascode = False
        try:
            # Get the selected item
            if selected_item is None:
                selected_item = self.doc.CurrentSelection

            if not hasattr(selected_item, 'supportsService'):
                self.msgbox(self.strings["errsel1"])
                logger.debug("Invalid selection (1)")
                return

            # cancel the use of character styles if context is not relevant
            if not self.charstylesavailable:
                self.options["UseCharStyles"] = False

            # TEXT SHAPE
            if selected_item.ImplementationName in ("SwXShape", "SvxShapeText", "com.sun.star.comp.sc.ScShapeObj"):
                logger.debug("Dealing with text shape.")
                code_block = selected_item
                code = code_block.String
                if code.strip():
                    hascode = True
                    lexer = self.getlexer(code_block)
                    # exit edit mode if necessary
                    self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                    undoaction = UndoAction(self.doc, code_block,
                                            f"code highlight (lang: {lexer.name}, style: {stylename})")
                    logger.debug("Custom undo action created.")
                    self.show_line_numbers(code_block, False)
                    code = code_block.String    # code string may have changed
                    cursor = code_block.createTextCursorByRange(code_block)
                    cursor.CharLocale = self.nolocale
                    self.highlight_code(cursor, lexer, style)
                    # unlock controllers here to force left pane syncing in draw/impress
                    if self.doc.supportsService("com.sun.star.drawing.GenericDrawingDocument"):
                        self.doc.unlockControllers()
                        logger.debug("Controllers unlocked.")
                    # code_block.FillStyle = FS_NONE
                    if bg_color:
                        code_block.FillStyle = FS_SOLID
                        code_block.FillColor = self.to_int(bg_color)
                    if self.options['ShowLineNumbers']:
                        self.show_line_numbers(code_block, True, charcolor=lineno_color)
                    # save options as user defined attribute
                    self.tagcodeblock(code_block, lexer.name)
                    # model is not considered as modified after textbox formatting
                    self.doc.setModified(True)
                    undoaction.get_new_state()
                    undomanager.addUndoAction(undoaction)
                    logger.debug("Custom undo action added.")

            # TEXT SHAPES
            elif selected_item.ImplementationName == "com.sun.star.drawing.SvxShapeCollection":
                logger.debug("Dealing with text shapes.")
                for code_block in selected_item:
                    code = code_block.String
                    if code.strip():
                        if updatecode:
                            udas = code_block.UserDefinedAttributes
                            if udas and SNIPPETTAGID in udas:
                                options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                                self.options.update(options)
                            else:
                                continue
                        hascode = True
                        self.prepare_highlight(code_block)

            # PLAIN TEXT
            elif selected_item.ImplementationName == "SwXTextRange":
                logger.debug("Dealing with plain text.")
                code_block = selected_item
                code = code_block.String
                if code.strip():
                    hascode = True
                    lexer = self.getlexer(code_block)
                    try:
                        undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                        self.show_line_numbers(code_block, False, isplaintext=True)
                        cursor, code = self.ensure_paragraphs(code_block)
                        # ParaBackColor does not work anymore, and new FillProperties isn't available from API
                        # see https://bugs.documentfoundation.org/show_bug.cgi?id=99125
                        # so let's use the dispatcher as workaround
                        # cursor.ParaBackColor = -1
                        # prop = PropertyValue(Name="BackgroundColor", Value=-1)
                        # self.dispatcher.executeDispatch(self.frame, ".uno:BackgroundColor", "", 0, (prop,))
                        self.doc.CurrentController.select(cursor)
                        cursor.CharLocale = self.nolocale
                        if bg_color and not self.inlinesnippet:
                            # cursor.ParaBackColor = self.to_int(bg_color)
                            prop = PropertyValue(Name="BackgroundColor", Value=self.to_int(bg_color))
                            self.dispatcher.executeDispatch(self.frame, ".uno:BackgroundColor", "", 0, (prop,))
                        self.highlight_code(cursor, lexer, style, bg_color)
                        if self.options['ShowLineNumbers']:
                            self.show_line_numbers(code_block, True, charcolor=lineno_color, isplaintext=True)
                        # save options as user defined attribute
                        self.tagcodeblock(code_block, lexer.name)
                    finally:
                        undomanager.leaveUndoContext()

            # PLAIN TEXTS
            elif selected_item.ImplementationName == "SwXTextRanges":
                for code_block in selected_item:
                    self.checkinlinesnippet(code_block)
                    code = code_block.String
                    if code.strip():
                        if updatecode:
                            if self.inlinesnippet:
                                udas = code_block.TextUserDefinedAttributes
                            else:
                                udas = code_block.ParaUserDefinedAttributes
                            if udas and SNIPPETTAGID in udas:
                                hascode = True
                                options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                                self.options.update(options)
                            else:
                                continue
                        hascode = True
                        self.prepare_highlight(code_block)

                if not hascode and selected_item.Count == 1:
                    code_block = selected_item[0]
                    if code_block.TextFrame:
                        self.prepare_highlight(code_block.TextFrame, updatecode)
                        return
                    elif code_block.TextTable:
                        cellname = code_block.Cell.CellName
                        texttablecursor = code_block.TextTable.createCursorByCellName(cellname)
                        self.prepare_highlight(texttablecursor, updatecode)
                        return
                    else:
                        if self.inlinesnippet:
                            udas = code_block.TextUserDefinedAttributes
                        else:
                            udas = code_block.ParaUserDefinedAttributes
                        if udas and SNIPPETTAGID in udas:
                            cursor, _ = self.ensure_paragraphs(code_block)
                            if updatecode:
                                options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                                self.options.update(options)
                            self.doc.CurrentController.select(cursor)
                            self.prepare_highlight(updatecode=updatecode)
                            return

            # TEXT FRAME
            elif selected_item.ImplementationName == "SwXTextFrame":
                if updatecode:
                    # Frame's UserDefinedAttributes can't be reverted with undo manager -> using text cursor instead
                    c = selected_item.createTextCursorByRange(selected_item)
                    udas = c.ParaUserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                        self.options.update(options)
                        hascode = True
                        self.prepare_highlight(selected_item)
                else:
                    logger.debug("Dealing with a text frame")
                    code_block = selected_item
                    code = code_block.String
                    if code.strip():
                        hascode = True
                        cursor = code_block.createTextCursorByRange(code_block)
                        lexer = self.getlexer(cursor)
                        undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                        self.show_line_numbers(code_block, False)
                        code = code_block.String    # code string may have changed
                        cursor = code_block.createTextCursorByRange(code_block)
                        try:
                            # code_block.BackColor = -1
                            if bg_color:
                                code_block.BackColor = self.to_int(bg_color)
                            cursor.CharLocale = self.nolocale
                            self.highlight_code(cursor, lexer, style)
                            if self.options['ShowLineNumbers']:
                                self.show_line_numbers(code_block, True, charcolor=lineno_color)
                            # save options as user defined attribute
                            cursor = code_block.createTextCursorByRange(code_block)
                            self.tagcodeblock(cursor, lexer.name)
                        finally:
                            undomanager.leaveUndoContext()

            # TEXT TABLE CELL
            elif selected_item.ImplementationName == "SwXCell":
                code_block = selected_item
                code = code_block.String
                if code.strip():
                    hascode = True
                    lexer = self.getlexer(code_block)
                    undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    self.show_line_numbers(code_block, False)
                    code = code_block.String    # code string may have changed
                    try:
                        # code_block.BackColor = -1
                        if bg_color:
                            code_block.BackColor = self.to_int(bg_color)
                        cursor = code_block.createTextCursorByRange(code_block)
                        cursor.CharLocale = self.nolocale
                        self.highlight_code(cursor, lexer, style)
                        if self.options['ShowLineNumbers']:
                            self.show_line_numbers(code_block, True, charcolor=lineno_color)
                        # save options as user defined attribute
                        self.tagcodeblock(code_block, lexer.name)
                    finally:
                        undomanager.leaveUndoContext()

            # TEXT TABLE CELL RANGE
            elif selected_item.ImplementationName == "SwXTextTableCursor":
                table = self.doc.CurrentController.ViewCursor.TextTable
                rangename = selected_item.RangeName
                if ':' not in rangename:
                    # only one cell
                    logger.debug("Dealing with a single text table cell.")
                    code_block = table.getCellByName(rangename)
                    if updatecode:
                        udas = code_block.UserDefinedAttributes
                        if udas and SNIPPETTAGID in udas:
                            options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                            self.options.update(options)
                            hascode = True
                            self.prepare_highlight(code_block)
                    else:
                        code = code_block.String
                        if code.strip():
                            hascode = True
                            self.prepare_highlight(code_block)
                else:
                    # at least two cells
                    logger.debug("Dealing with multiple text table cells.")
                    cellrange = table.getCellRangeByName(rangename)
                    nrows, ncols = len(cellrange.Data), len(cellrange.Data[0])
                    for row in range(nrows):
                        for col in range(ncols):
                            code_block = cellrange.getCellByPosition(col, row)
                            if updatecode:
                                udas = code_block.UserDefinedAttributes
                                if udas and SNIPPETTAGID in udas:
                                    options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                                    self.options.update(options)
                                    hascode = True
                                    self.prepare_highlight(code_block)
                                    continue
                            else:
                                code = code_block.String
                                if code.strip():
                                    hascode = True
                                    self.prepare_highlight(code_block)

            # CURSOR INSIDE DRAW/IMPRESS SHAPE
            elif selected_item.ImplementationName == "SvxUnoTextCursor":
                logger.debug("Dealing with text shape in edit mode.")
                # exit edit mode
                self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                self.prepare_highlight(updatecode=updatecode)
                return

                # ### OLD CODE, intended to highlight sub text, but api's too buggy'
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

            # CALC CELL

            # CALC CELL RANGE
            elif selected_item.ImplementationName in ("ScCellObj", "ScCellRangeObj", "ScCellRangesObj"):
                logger.debug('Dealing with Calc cells.')
                # exit edit mode if necessary
                self.dispatcher.executeDispatch(self.frame, ".uno:Deselect", "", 0, ())
                cells = selected_item.queryContentCells(CF_STRING).Cells
                if cells.hasElements():
                    for code_block in cells:
                        if updatecode:
                            udas = code_block.UserDefinedAttributes
                            if udas and SNIPPETTAGID in udas:
                                hascode = True
                                options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                                self.options.update(options)
                                self.prepare_highlight(code_block)
                            else:
                                continue
                        else:
                            hascode = True
                            code = code_block.String
                            lexer = self.getlexer(code_block)
                            undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                            self.show_line_numbers(code_block, False)
                            code = code_block.String    # code string may have changed
                            try:
                                # code_block.CellBackColor = -1
                                code_block.CharLocale = self.nolocale
                                if bg_color:
                                    code_block.CellBackColor = self.to_int(bg_color)
                                cursor = code_block.createTextCursor()
                                self.highlight_code(cursor, lexer, style)
                                if self.options['ShowLineNumbers']:
                                    self.show_line_numbers(code_block, True, charcolor=lineno_color)
                                # save options as user defined attribute
                                self.tagcodeblock(code_block, lexer.name)
                            finally:
                                undomanager.leaveUndoContext()

            else:
                logger.debug("Invalid selection (2).")
                self.msgbox(self.strings["errsel1"])
                return

            if not hascode:
                if updatecode:
                    logger.debug("Selection is not updatable.")
                    self.msgbox(self.strings["errsel3"])
                else:
                    logger.debug("Current selection contains no text.")
                    self.msgbox(self.strings["errsel2"])

        except AttributeError:
            self.msgbox(self.strings["errsel1"])
            logger.exception("")
        except Exception:
            self.msgbox(traceback.format_exc())
            logger.exception("")
        finally:
            if self.doc.hasControllersLocked():
                self.doc.unlockControllers()
                logger.debug("Controllers unlocked.")

    def highlight_code(self, cursor, lexer, style, inline_bg_color=None):
        def _highlight_code():
            cursor.goRight(len(lastval), True)  # selects the token's text
            try:
                if self.options["UseCharStyles"]:
                    cursor.CharStyleName = str(lasttype).replace('Token', styleprefix)
                else:
                    tok_style = style.style_for_token(lasttype)
                    cursor.CharColor = self.to_int(tok_style['color'])
                    cursor.CharWeight = W_BOLD if tok_style['bold'] else W_NORMAL
                    cursor.CharPosture = SL_ITALIC if tok_style['italic'] else SL_NONE
                    cursor.CharUnderline = UL_SINGLE if tok_style['underline'] else UL_NONE
                    if tok_style["bgcolor"]:
                        cursor.CharBackColor = self.to_int(tok_style["bgcolor"])
                    elif self.inlinesnippet and inline_bg_color:
                        cursor.CharBackColor = self.to_int(inline_bg_color)
            except Exception:
                pass
            finally:
                cursor.collapseToEnd()  # deselects the selected text
        code = cursor.String
        # clean up any previous formatting
        cursor.setPropertyValues(("CharBackColor", "CharColor", "CharPosture", "CharUnderline", "CharWeight"),
                                 (-1, -1, SL_NONE, UL_NONE, W_NORMAL))
        if self.charstylesavailable:
            cursor.setPropertiesToDefault(("CharStyleName", "CharStyleNames"))
        cursor.collapseToStart()

        # create character styles if requested
        # (this happens here to stay synched with undo context)
        styleprefix = CHARSTYLEID + style.__name__.lower()[:-5]
        if self.options["UseCharStyles"]:
            self.createcharstyles(style, styleprefix)
        # caching consecutive tokens with same token type
        logger.debug(f"Starting code block highlighting (lexer: {lexer}, style: {style}).")
        lastval = ''
        lasttype = None
        for tok_type, tok_value in lexer.get_tokens(code):
            if tok_type == lasttype:
                lastval += tok_value
            else:
                if lastval:
                    _highlight_code()
                lastval = tok_value
                lasttype = tok_type
        # emptying buffer
        if lastval.strip():
            _highlight_code()
        self.cleancharstyles(styleprefix)
        logger.debug("Terminating code block highlighting.")

    def show_line_numbers(self, code_block, show, charcolor=-1, isplaintext=False):
        if self.inlinesnippet:
            return
        startnb = self.options["LineNumberStart"]
        ratio = self.options["LineNumberRatio"]
        sep = self.options["LineNumberSeparator"]
        pad = self.options["LineNumberPaddingSymbol"]
        logger.debug(f"Starting code block numbering (show: {show}).")
        sep = sep.replace(r'\t', '\t')
        codecharheight = code_block.End.CharHeight
        nocharheight = round(codecharheight*ratio//50)/2   # round to 0.5

        if isplaintext:
            c = code_block.Text.createTextCursorByRange(code_block)
            code = c.String
        else:
            c = code_block.Text.createTextCursor()
            code = c.Text.String

        def show_numbering():
            nblignes = len(code.split('\n'))
            digits = int(log10(nblignes - 1 + startnb)) + 1
            for n, para in enumerate(code_block, start=startnb):
                # para.Start.CharHeight = nocharheight
                prefix = f'{n:{pad}>{digits}}{sep}'
                para.Start.setString(prefix)
                c.gotoRange(para.Start, False)
                c.goRight(len(prefix), True)
                c.CharHeight = nocharheight
                c.setPropertyValues(("CharColor", "CharPosture", "CharUnderline", "CharWeight"),
                                    (charcolor, SL_NONE, UL_NONE, W_NORMAL))

        def hide_numbering():
            for para in code_block:
                if p.match(para.String):
                    para.CharHeight = codecharheight
                    para.String = para.String[lenno:]

        if show:
            logger.debug("Showing code block numbering.")
            show_numbering()
        else:
            # check for existing line numbering and its width
            p = re.compile(r"^\s*[0-9]+[\W]?\s+", re.MULTILINE)
            try:
                lenno = min(len(f) for f in p.findall(code))
            except ValueError:
                lenno = None
            if lenno:
                logger.debug("Hiding code block numbering.")
                hide_numbering()

    def checkinlinesnippet(self, code_block):
        self.inlinesnippet = False
        c = code_block.Text.createTextCursorByRange(code_block)
        if code_block.String:
            if c.Start.TextParagraph == c.End.TextParagraph:
                if c.Text.compareRegionStarts(c, c.Start.TextParagraph) != 0:
                    # inline snippet
                    logger.info('Code identified as inline snippet.')
                    self.inlinesnippet = True
        else:
            udas = code_block.TextUserDefinedAttributes
            if udas and SNIPPETTAGID in udas:
                self.inlinesnippet = True

    def ensure_paragraphs(self, selected_code):
        '''Ensure the selection does not contains part of paragraphs.
        Cursor could start or end in the middle of a code line, when plain text selected.
        So let's expand it to the entire paragraphs.'''

        c = selected_code.Text.createTextCursorByRange(selected_code)
        if selected_code.String:
            if self.inlinesnippet:
                # inline snippet, abort expansion
                return c, c.String
            c.gotoStartOfParagraph(False)
            c.gotoRange(selected_code.End, True)
            c.gotoEndOfParagraph(True)
        else:
            if self.inlinesnippet:
                udas = selected_code.TextUserDefinedAttributes
                options = udas.getByName(SNIPPETTAGID).Value
                while c.goLeft(1, False):
                    udas2 = c.TextUserDefinedAttributes
                    if not (udas2 and SNIPPETTAGID in udas2 and udas2.getByName(SNIPPETTAGID).Value == options):
                        break

                c.gotoRange(selected_code, True)
                while c.goRight(1, True):
                    udas2 = c.TextUserDefinedAttributes
                    if not (udas2 and SNIPPETTAGID in udas2 and udas2.getByName(SNIPPETTAGID).Value == options):
                        c.goLeft(1, True)
                        break
            else:
                udas = selected_code.ParaUserDefinedAttributes
                options = udas.getByName(SNIPPETTAGID).Value
                startpara = c.TextParagraph
                endpara = c.TextParagraph
                while c.gotoPreviousParagraph(False):
                    udas2 = c.ParaUserDefinedAttributes
                    if (udas2 and SNIPPETTAGID in udas2 and udas2.getByName(SNIPPETTAGID).Value == options):
                        startpara = c.TextParagraph
                    else:
                        break

                c.gotoRange(selected_code, False)
                while c.gotoNextParagraph(False):
                    udas2 = c.ParaUserDefinedAttributes
                    if (udas2 and SNIPPETTAGID in udas2 and udas2.getByName(SNIPPETTAGID).Value == options):
                        endpara = c.TextParagraph
                    else:
                        break
                c.gotoRange(startpara.Start, False)
                c.gotoRange(endpara.End, True)
        return c, c.String

    # dev tools
    def update_all(self, usetags):
        '''Update all formatted code in the active document.
        DO NOT PUBLISH, ALPHA VERSION'''
        def highlight_snippet(code_block, udas):
            if usetags:
                options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                self.options.update(options)
            self.doc.CurrentController.select(code_block)
            logger.debug(f'Updating snippet (type: {code_block.ImplementationName})')
            self.prepare_highlight(self.doc.CurrentSelection)

        def browsetaggedcode_text(container=None):
            root = False
            self.charstylesavailable = True
            if not container:
                container = self.doc.Text
                root = True
            cursor = None
            options = None
            for para in container:
                if not para.supportsService('com.sun.star.text.Paragraph'):
                    continue
                udas = para.ParaUserDefinedAttributes
                if udas is None:
                    continue
                if SNIPPETTAGID in udas and not cursor:
                    options = udas
                    cursor = container.createTextCursorByRange(para.Start)
                elif cursor and SNIPPETTAGID not in udas:
                    cursor.gotoRange(para.Start, True)
                    cursor.goLeft(1, True)
                    highlight_snippet(cursor, options)
                    cursor, options = None, None
            # last paragraph could be part of a code block
            if cursor and SNIPPETTAGID in udas:
                cursor.gotoRange(para.End, True)
                highlight_snippet(cursor, udas)

            if root:
                for frame in self.doc.TextFrames:
                    self.charstylesavailable = True
                    udas = frame.UserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        highlight_snippet(frame, udas)
                    else:
                        browsetaggedcode_text(frame)
                for table in self.doc.TextTables:
                    self.doc.CurrentController.select(table)
                    self.charstylesavailable = True
                    cellnames = table.CellNames
                    for cellname in cellnames:
                        cell = table.getCellByName(cellname)
                        udas = cell.UserDefinedAttributes
                        if udas and SNIPPETTAGID in udas:
                            highlight_snippet(cell, udas)
                        else:
                            browsetaggedcode_text(cell)
                for shape in self.doc.DrawPage:
                    if shape.ImplementationName == "SwXShape":
                        udas = shape.UserDefinedAttributes
                        if udas and SNIPPETTAGID in udas:
                            self.charstylesavailable = False
                            highlight_snippet(shape, udas)

        def browsetaggedcode_calc():
            for sheet in self.doc.Sheets:
                for ranges in sheet.UniqueCellFormatRanges:
                    udas = ranges.UserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        highlight_snippet(ranges, udas)

        def browsetaggedcode_draw():
            for drawpage in self.doc.DrawPages:
                for shape in drawpage:
                    try:
                        udas = shape.UserDefinedAttributes
                        if udas and SNIPPETTAGID in udas:
                            self.charstylesavailable = False
                            highlight_snippet(shape, udas)
                    except AttributeError:
                        continue

        logger.debug("Updating all snippets previously formatted with Code Highlighter 2.")
        sel = self.doc.CurrentSelection
        try:
            if self.doc.supportsService('com.sun.star.text.GenericTextDocument'):
                browsetaggedcode_text()
            elif self.doc.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                browsetaggedcode_calc()
            elif self.doc.supportsService('com.sun.star.drawing.GenericDrawingDocument'):
                browsetaggedcode_draw()
            self.msgbox("Done.")
        finally:
            self.doc.CurrentController.select(sel)

    def removealltags(self):
        '''
        Remove all snippet tags in the active document.
        TODO:
        - progress bar
        - selection only
        - draw and impress: add custom undo action
        '''
        def searchforlexertag_text(container=None):
            root = False
            if not container:
                container = self.doc.Text
                root = True
            c = container.Text.createTextCursorByRange(container)
            # search for ParaUserDefinedAttributes
            udas = c.ParaUserDefinedAttributes
            if not udas:    # ParaUserDefinedAttributes is empty when container contains mixed attributes
                for para in container:
                    if not para.supportsService('com.sun.star.text.Paragraph'):
                        continue
                    udas2 = para.ParaUserDefinedAttributes
                    if SNIPPETTAGID in udas2:
                        udas2.removeByName(SNIPPETTAGID)
                        para.ParaUserDefinedAttributes = udas2
            elif SNIPPETTAGID in udas:
                udas.removeByName(SNIPPETTAGID)
                c.ParaUserDefinedAttributes = udas

            # search for TextUserDefinedAttributes
            udas = c.TextUserDefinedAttributes
            if not udas:    # TextUserDefinedAttributes is empty when container contains mixed attributes
                for para in container:
                    if not para.supportsService('com.sun.star.text.Paragraph'):
                        continue
                    udas2 = para.TextUserDefinedAttributes
                    if not udas2:
                        for portion in para:
                            if portion.TextPortionType == "Text":
                                udas3 = portion.TextUserDefinedAttributes
                                if SNIPPETTAGID in udas3:
                                    udas3.removeByName(SNIPPETTAGID)
                                    portion.TextUserDefinedAttributes = udas3
                    elif SNIPPETTAGID in udas2:
                        udas2.removeByName(SNIPPETTAGID)
                        para.TextUserDefinedAttributes = udas2
            elif SNIPPETTAGID in udas:
                udas.removeByName(SNIPPETTAGID)
                c.TextUserDefinedAttributes = udas

            # search for UserDefinedAttributes and for subtexts
            if root:
                for frame in self.doc.TextFrames:
                    udas = frame.UserDefinedAttributes
                    if SNIPPETTAGID in udas:
                        udas.removeByName(SNIPPETTAGID)
                        frame.UserDefinedAttributes = udas
                    searchforlexertag_text(frame)
                for table in self.doc.TextTables:
                    cellnames = table.CellNames
                    for cellname in cellnames:
                        cell = table.getCellByName(cellname)
                        udas = cell.UserDefinedAttributes
                        if SNIPPETTAGID in udas:
                            udas.removeByName(SNIPPETTAGID)
                            cell.UserDefinedAttributes = udas
                        searchforlexertag_text(cell)
                for shape in self.doc.DrawPage:
                    udas = shape.UserDefinedAttributes
                    if SNIPPETTAGID in udas:
                        udas.removeByName(SNIPPETTAGID)
                        shape.UserDefinedAttributes = udas

        def searchforlexertag_calc():
            for sheet in self.doc.Sheets:
                for ranges in sheet.UniqueCellFormatRanges:
                    udas = ranges.UserDefinedAttributes
                    if SNIPPETTAGID in udas:
                        udas.removeByName(SNIPPETTAGID)
                        ranges.UserDefinedAttributes = udas

        def searchforlexertag_draw():
            for drawpage in self.doc.DrawPages:
                for shape in drawpage:
                    try:
                        udas = shape.UserDefinedAttributes
                        if SNIPPETTAGID in udas:
                            udas.removeByName(SNIPPETTAGID)
                            shape.UserDefinedAttributes = udas
                    except AttributeError:
                        continue

        self.doc.lockControllers()
        undomanager = self.doc.UndoManager
        undomanager.enterUndoContext("All CH2 attributes removed.")
        try:
            if self.doc.supportsService('com.sun.star.text.GenericTextDocument'):
                searchforlexertag_text()
                self.msgbox("Done.")
            elif self.doc.supportsService('com.sun.star.sheet.SpreadsheetDocument'):
                searchforlexertag_calc()
                self.msgbox("Done.")
            elif self.doc.supportsService('com.sun.star.drawing.GenericDrawingDocument'):
                searchforlexertag_draw()
                self.msgbox("Done.")
            else:
                self.msgbox("Module not yet supported.")
        finally:
            undomanager.leaveUndoContext()
            self.doc.unlockControllers()


# Component registration
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(CodeHighlighter, "ooo.ext.code-highlighter.impl", (),)


# exposed functions for development stages only
# uncomment corresponding entry in ../META_INF/manifest.xml to add them as framework scripts
def highlight(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_highlight()


def highlight_previous(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_highlight_previous()


def highlight_update(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_update()


def highlight_update_all(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.update_all(True)

def remove_all_tags(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.do_removealltags()
