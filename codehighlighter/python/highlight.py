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
import sys
import uno
import os.path
import logging
from com.sun.star.uno import RuntimeException

try:
    LOGLEVEL = {0: logging.WARNING, 1: logging.INFO, 2: logging.DEBUG}
    logger = logging.getLogger("codehighlighter")
    formatter = logging.Formatter("%(levelname)s [%(funcName)s::%(lineno)d] %(message)s")
    logger.handlers[:] = []
    consolehandler = logging.StreamHandler()
    consolehandler.setFormatter(formatter)
    logger.addHandler(consolehandler)
    logger.setLevel(logging.INFO)
    logger.info("Logger installed.")
    userpath = uno.getComponentContext().ServiceManager.createInstance(
                    "com.sun.star.util.PathSubstitution").substituteVariables("$(user)", True)
    logfile = os.path.join(uno.fileUrlToSystemPath(userpath), "codehighlighter.log")
    filehandler = logging.FileHandler(logfile, mode="w", delay=True)
    filehandler.setFormatter(formatter)
except RuntimeException:
    # At installation time, no context is available -> just ignore it.
    pass


# simple import hook, making sure embedded pygments is found first
try:
    path = os.path.join(os.path.dirname(__file__), "pythonpath")
    sys.path.insert(0, sys.path.pop(sys.path.index(path)))
    logger.debug(f'sys.path: {sys.path}')
    logger.info("Embedded Pygments path priorised.")
except NameError:
    # __file__ is not defined
    # only occurs when using exposed functions -> should be harmless
    pass
except Exception:
    logger.exception("")

# other imports
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
INVALID_SELECTION = "Invalid"


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
        self.charprops = ("CharBackColor", "CharColor", "CharLocale", "CharPosture",
                          "CharHeight", "CharUnderline", "CharWeight")
        self.bgprops = ("FillColor", "FillStyle")
        self.len_ = self.define_len()
        self.get_old_state()
        # XUndoAction attribute
        self.Title = title

    def define_len(self):
        # workaround for issue 22 (https://github.com/jmzambon/libreoffice-code-highlighter/issues/22)
        if any(ord(char) >= 0x10000 for char in self.textbox.String):
            return lambda s: sum(1 if ord(char) < 0x10000 else 2 for char in s)
        return len

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
                plen = self.len_(portion.String)
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
            self.undomanager = self.doc.UndoManager
            self.selection = None
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
            self.activepreviews = 0
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
        elif method == "preview":
            focus = {1: 'cb_style', 2: 'nb_start'}
            self.do_preview(dialog)
            dialog.getControl(focus[dialog.Model.Step]).setFocus()
            return True
        return False

    def getSupportedMethodNames(self):
        return 'preview', 'topage1', 'topage2'

    # main functions
    def do_highlight(self):
        '''Open option dialog and start code highlighting.'''

        self.selection = self.check_selection()
        if self.selection == INVALID_SELECTION:
            self.msgbox(self.strings["errsel1"])
        elif self.selection:
            ret = self.choose_options()
            logger.debug("Undoing existing previews on dialog closing.")
            while self.activepreviews:
                self.undomanager.undo()
                self.undomanager.clearRedo()
                self.activepreviews -= 1
            if ret:
                logger.debug("Starting highlights.")
                for code_block in self.selection:
                    self.prepare_highlight(code_block)
        else:
            logger.debug("Current selection contains no text.")
            self.msgbox(self.strings["errsel2"])

    def do_highlight_previous(self):
        '''Start code highlighting with current options as default.'''

        selection = self.check_selection()
        if selection == INVALID_SELECTION:
            self.msgbox(self.strings["errsel1"])
        elif selection:
            for code_block in selection:
                self.prepare_highlight(code_block)
        else:
            logger.debug("Current selection contains no text.")
            self.msgbox(self.strings["errsel2"])

    def do_update(self):
        '''Update already highlighted snippets based on options stored in codeblock tags.
        Code-blocks must have been highlighted at least once with Code Highlighter 2.'''

        hasupdates = False
        selection = self.check_selection()
        if selection == INVALID_SELECTION:
            self.msgbox(self.strings["errsel1"])
        elif selection:
            for code_block in selection:
                ret = self.prepare_highlight(code_block, updatecode=True)
                if not hasupdates:
                    hasupdates = ret
            if not hasupdates:
                logger.debug("Selection is not updatable.")
                self.msgbox(self.strings["errsel3"])
        else:
            logger.debug("Current selection contains no text.")
            self.msgbox(self.strings["errsel2"])

    def do_preview(self, dialog):
        logger.debug("Undoing existing previews before creating new ones.")
        while self.activepreviews:
            self.undomanager.undo()
            self.activepreviews -= 1
        choices = self.get_options_from_dialog(dialog)
        if choices:
            logger.debug("Creating previews.")
            self.options.update(choices)
            for code_block in self.selection:
                self.prepare_highlight(code_block)
                self.activepreviews += 1

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
            for h in logger.handlers:
                if isinstance(h, logging.FileHandler):
                    logger.removeHandler(h)
                    return
        else:
            for h in logger.handlers:
                if isinstance(h, logging.FileHandler):
                    return
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
                        "lbl_cs_rootstyle", "pygments_ver", "preview", "topage1", "topage2")
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

    def get_options_from_dialog(self, dialog):
        opt = {}
        lang = dialog.getControl('cb_lang').Text.strip() or 'automatic'
        style = dialog.getControl('cb_style').Text.strip() or 'default'
        if lang != 'automatic' and lang.lower() not in self.all_lexer_aliases:
            self.msgbox(self.strings["errlang"])
        elif style not in self.all_styles:
            self.msgbox(self.strings["errstyle"])
        else:
            opt['Language'] = lang
            opt['Style'] = style
            opt['ColourizeBackground'] = dialog.getControl('check_col_bg').State
            opt['UseCharStyles'] = dialog.getControl('check_charstyles').State
            opt['ShowLineNumbers'] = dialog.getControl('check_linenb').State
            opt['LineNumberStart'] = int(dialog.getControl('nb_start').Value)
            opt['LineNumberRatio'] = int(dialog.getControl('nb_ratio').Value)
            opt['LineNumberSeparator'] = dialog.getControl('nb_sep').Text
            opt['LineNumberPaddingSymbol'] = dialog.getControl('nb_pad').Text
            opt['MasterCharStyle'] = dialog.getControl('cs_rootstyle').Text
        return opt

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

        choices = self.get_options_from_dialog(dialog)
        if not choices:
            return False

        self.save_options(choices)
        logger.debug("Dialog validated and options saved.")
        logger.info(f"Updated options = {self.options}.")
        return True

    def save_options(self, choices):
        self.options.update(choices)
        self.cfg_access.setPropertyValues(tuple(choices.keys()), tuple(choices.values()))
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

    def check_selection(self):
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
        # Get the selected item
        selected_item = self.doc.CurrentSelection

        if not hasattr(selected_item, 'supportsService'):
            logger.debug("Invalid selection (1)")
            return INVALID_SELECTION

        code_blocks = []

        # TEXT SHAPES
        if selected_item.ImplementationName == "com.sun.star.drawing.SvxShapeCollection":
            logger.debug("Checking selection: com.sun.star.drawing.SvxShapeCollection.")
            for code_block in selected_item:
                if code_block.String.strip():
                    # exit edit mode if necessary
                    self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
                    try:
                        if code_block.TextBox:
                            code_blocks.append(code_block.TextBoxContent)
                        else:
                            code_blocks.append(code_block)
                    except AttributeError:
                        code_blocks.append(code_block)

        # PLAIN TEXTS
        elif selected_item.ImplementationName == "SwXTextRanges":
            logger.debug("Checking selection: SwXTextRanges.")
            for code_block in selected_item:
                if code_block.String.strip():
                    code_blocks.append(code_block)
                elif selected_item.Count == 1:
                    code_block = selected_item[0]
                    if code_block.TextFrame and code_block.Text.String.strip():
                        code_blocks.append(code_block.TextFrame)
                    elif code_block.TextTable and code_block.Text.String.strip():
                        cellname = code_block.Cell.CellName
                        cell = code_block.TextTable.getCellByName(cellname)
                        code_blocks.append(cell)
                    else:
                        self.checkinlinesnippet(code_block)
                        if self.inlinesnippet:
                            udas = code_block.TextUserDefinedAttributes
                        else:
                            udas = code_block.ParaUserDefinedAttributes
                        if udas and SNIPPETTAGID in udas:
                            cursor, _ = self.ensure_paragraphs(code_block)
                            self.doc.CurrentController.select(cursor)
                            code_blocks.append(self.doc.CurrentSelection[0])

        # TEXT FRAME
        elif selected_item.ImplementationName == "SwXTextFrame":
            logger.debug("Checking selection: SwXTextFrame.")
            if selected_item.String.strip():
                code_blocks.append(selected_item)

        # TEXT TABLE CELL RANGE
        elif selected_item.ImplementationName == "SwXTextTableCursor":
            logger.debug("Checking selection: SwXTextTableCursor.")
            table = self.doc.CurrentController.ViewCursor.TextTable
            rangename = selected_item.RangeName
            if ':' not in rangename:
                # only one cell
                code_block = table.getCellByName(rangename)
                if code_block.String.strip():
                    code_blocks.append(code_block)
            else:
                # at least two cells
                cellrange = table.getCellRangeByName(rangename)
                nrows, ncols = len(cellrange.Data), len(cellrange.Data[0])
                for row in range(nrows):
                    for col in range(ncols):
                        code_block = cellrange.getCellByPosition(col, row)
                        if code_block.String.strip():
                            code_blocks.append(code_block)

        # CALC CELL RANGE
        elif selected_item.ImplementationName in ("ScCellObj", "ScCellRangeObj", "ScCellRangesObj"):
            logger.debug("Checking selection: calc cell, cell range or cell ranges.")
            # exit edit mode if necessary
            self.dispatcher.executeDispatch(self.frame, ".uno:Deselect", "", 0, ())
            cells = selected_item.queryContentCells(CF_STRING).Cells
            for cell in cells:
                code_blocks.append(cell)

        # CURSOR INSIDE DRAW/IMPRESS SHAPE
        elif selected_item.ImplementationName == "SvxUnoTextCursor":
            logger.debug("Checking selection: SvxUnoTextCursor (text inside draw/impress shape).")
            # exit edit mode
            self.dispatcher.executeDispatch(self.frame, ".uno:SelectObject", "", 0, ())
            selected_item = self.doc.CurrentSelection
            for code_block in selected_item:
                if code_block.String.strip():
                    code_blocks.append(code_block)

        else:
            logger.debug("Invalid selection (2).")
            return INVALID_SELECTION

        return code_blocks

    def prepare_highlight(self, code_block, updatecode=False):
        if not self.doc.hasControllersLocked():
            self.doc.lockControllers()
            logger.debug("Controllers locked.")

        controller = self.doc.CurrentController
        hascode = True
        try:
            # cancel the use of character styles if context is not relevant
            if not self.charstylesavailable:
                self.options["UseCharStyles"] = False

            # TEXT SHAPE
            if code_block.ImplementationName in ("SwXShape", "SvxShapeText", "com.sun.star.comp.sc.ScShapeObj"):
                logger.debug("Dealing with text shape.")

                if updatecode:
                    udas = code_block.UserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                        self.options.update(options)
                    else:
                        hascode = False

                if hascode:
                    stylename = self.options['Style']
                    style = self.getstylebyname(stylename)
                    bg_color = style.background_color if self.options['ColourizeBackground'] else None
                    lineno_color = -1
                    if self.options['ShowLineNumbers']:
                        _style = style.style_for_token(("Comment",))
                        lineno_color = self.to_int(_style['color'])
                    lexer = self.getlexer(code_block)

                    undoaction = UndoAction(self.doc, code_block,
                                            f"code highlight (lang: {lexer.name}, style: {stylename})")
                    logger.debug("Custom undo action created.")
                    self.show_line_numbers(code_block, False)
                    cursor = code_block.createTextCursorByRange(code_block)
                    cursor.CharLocale = self.nolocale
                    self.highlight_code(cursor, lexer, style, checkunicode=True)
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
                    self.undomanager.addUndoAction(undoaction)
                    logger.debug("Custom undo action added.")

            # PLAIN TEXT
            elif code_block.ImplementationName == "SwXTextRange":
                logger.debug("Dealing with plain text.")
                self.checkinlinesnippet(code_block)

                if updatecode:
                    if self.inlinesnippet:
                        udas = code_block.TextUserDefinedAttributes
                    else:
                        udas = code_block.ParaUserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                        self.options.update(options)
                        cursor, _ = self.ensure_paragraphs(code_block)
                        self.doc.CurrentController.select(cursor)
                        code_block = self.doc.CurrentSelection[0]
                    else:
                        hascode = False

                if hascode:
                    stylename = self.options['Style']
                    style = self.getstylebyname(stylename)
                    bg_color = style.background_color if self.options['ColourizeBackground'] else None
                    lineno_color = -1
                    if self.options['ShowLineNumbers']:
                        _style = style.style_for_token(("Comment",))
                        lineno_color = self.to_int(_style['color'])

                    try:
                        self.show_line_numbers(code_block, False, isplaintext=True)
                        cursor, code = self.ensure_paragraphs(code_block)
                        lexer = self.getlexer(cursor)
                        self.undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                        # ParaBackColor does not work anymore, and new FillProperties isn't available from API
                        # see https://bugs.documentfoundation.org/show_bug.cgi?id=99125
                        # so let's use the dispatcher as workaround
                        # cursor.ParaBackColor = -1
                        # prop = PropertyValue(Name="BackgroundColor", Value=-1)
                        # self.dispatcher.executeDispatch(self.frame, ".uno:BackgroundColor", "", 0, (prop,))
                        controller.select(cursor)
                        cursor.CharLocale = self.nolocale
                        char_bg_color = None
                        if bg_color and not self.inlinesnippet:
                            # cursor.ParaBackColor = self.to_int(bg_color)
                            prop = PropertyValue(Name="BackgroundColor", Value=self.to_int(bg_color))
                            self.dispatcher.executeDispatch(self.frame, ".uno:BackgroundColor", "", 0, (prop,))
                        elif self.inlinesnippet:
                            char_bg_color = bg_color
                        self.highlight_code(cursor, lexer, style, char_bg_color=char_bg_color)
                        if self.options['ShowLineNumbers']:
                            self.show_line_numbers(code_block, True, charcolor=lineno_color, isplaintext=True)
                        # save options as user defined attribute
                        self.tagcodeblock(code_block, lexer.name)
                        controller.ViewCursor.collapseToEnd()
                    finally:
                        self.undomanager.leaveUndoContext()

            # TEXT FRAME
            elif code_block.ImplementationName == "SwXTextFrame":
                logger.debug("Dealing with a text frame")
                code_block = code_block

                if updatecode:
                    # Frame's UserDefinedAttributes can't be reverted with undo manager -> using text cursor instead
                    c = code_block.createTextCursorByRange(code_block)
                    udas = c.ParaUserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                        self.options.update(options)
                    else:
                        hascode = False

                if hascode:
                    stylename = self.options['Style']
                    style = self.getstylebyname(stylename)
                    bg_color = style.background_color if self.options['ColourizeBackground'] else None
                    lineno_color = -1
                    if self.options['ShowLineNumbers']:
                        _style = style.style_for_token(("Comment",))
                        lineno_color = self.to_int(_style['color'])
                    lexer = self.getlexer(code_block)

                    hascode = True
                    cursor = code_block.createTextCursorByRange(code_block)
                    # lexer = self.getlexer(cursor)
                    self.undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    self.show_line_numbers(code_block, False)
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
                        controller.select(code_block)
                    finally:
                        self.undomanager.leaveUndoContext()

            # TEXT TABLE CELL
            elif code_block.ImplementationName == "SwXCell":
                logger.debug("Dealing with a text table cell.")
                code_block = code_block
                hascode = True
                if updatecode:
                    udas = code_block.UserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                        self.options.update(options)
                    else:
                        hascode = False

                if hascode:
                    stylename = self.options['Style']
                    style = self.getstylebyname(stylename)
                    bg_color = style.background_color if self.options['ColourizeBackground'] else None
                    lineno_color = -1
                    if self.options['ShowLineNumbers']:
                        _style = style.style_for_token(("Comment",))
                        lineno_color = self.to_int(_style['color'])
                    lexer = self.getlexer(code_block)

                    self.undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    self.show_line_numbers(code_block, False)
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
                        controller.ViewCursor.collapseToEnd()
                    finally:
                        self.undomanager.leaveUndoContext()

            # CALC CELL
            # see this for info: https://bugs.documentfoundation.org/show_bug.cgi?id=151839
            elif code_block.ImplementationName in ("ScCellObj"):
                logger.debug('Dealing with Calc cell.')
                hascode = True
                code_block = code_block
                if updatecode:
                    udas = code_block.UserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                        self.options.update(options)
                    else:
                        hascode = False

                if hascode:
                    stylename = self.options['Style']
                    style = self.getstylebyname(stylename)
                    bg_color = style.background_color if self.options['ColourizeBackground'] else None
                    lineno_color = -1
                    if self.options['ShowLineNumbers']:
                        _style = style.style_for_token(("Comment",))
                        lineno_color = self.to_int(_style['color'])
                    lexer = self.getlexer(code_block)

                    self.undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                    self.show_line_numbers(code_block, False)
                    try:
                        # code_block.CellBackColor = -1
                        code_block.CharLocale = self.nolocale
                        if bg_color:
                            code_block.CellBackColor = self.to_int(bg_color)
                        cursor = code_block.createTextCursor()
                        self.highlight_code(cursor, lexer, style, char_bg_color=bg_color, checkunicode=True)
                        if self.options['ShowLineNumbers']:
                            self.show_line_numbers(code_block, True, charcolor=lineno_color, char_bg_color=bg_color)
                        # save options as user defined attribute
                        self.tagcodeblock(code_block, lexer.name)
                    finally:
                        self.undomanager.leaveUndoContext()

            return hascode

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

    def highlight_code(self, cursor, lexer, style, char_bg_color=None, checkunicode=False):
        def _highlight_code():
            cursor.goRight(len_(lastval), True)  # selects the token's text
            try:
                if self.options["UseCharStyles"]:
                    cursor.CharStyleName = str(lasttype).replace('Token', styleprefix)
                else:
                    tok_style = style.style_for_token(lasttype)
                    properties = ['CharColor', 'CharWeight', 'CharPosture', 'CharUnderline']
                    values = [self.to_int(tok_style['color']),
                              W_BOLD if tok_style['bold'] else W_NORMAL,
                              SL_ITALIC if tok_style['italic'] else SL_NONE,
                              UL_SINGLE if tok_style['underline'] else UL_NONE]
                    if tok_style["bgcolor"]:
                        properties.append('CharBackColor')
                        values.append(self.to_int(tok_style["bgcolor"]))
                    elif char_bg_color:
                        properties.append('CharBackColor')
                        values.append(self.to_int(char_bg_color))
                    cursor.setPropertyValues(properties, values)
            except Exception:
                pass
            finally:
                cursor.collapseToEnd()  # deselects the selected text

        code = cursor.String

        # workaround issue 22 (https://github.com/jmzambon/libreoffice-code-highlighter/issues/22)
        len_ = len
        if checkunicode and any(ord(char) >= 0x10000 for char in code):
            len_ = lambda s: sum(1 if ord(char) < 0x10000 else 2 for char in s)

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

    def show_line_numbers(self, code_block, show, charcolor=-1, isplaintext=False, char_bg_color=None):
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
            nblines = len(code.split('\n'))
            digits = int(log10(nblines - 1 + startnb)) + 1
            for n, para in enumerate(code_block, start=startnb):
                # para.Start.CharHeight = nocharheight
                prefix = f'{n:{pad}>{digits}}{sep}'
                para.Start.setString(prefix)
                c.gotoRange(para.Start, False)
                c.goRight(len(prefix), True)
                c.CharHeight = nocharheight
                c.setPropertyValues(("CharBackColor", "CharColor", "CharPosture", "CharUnderline", "CharWeight"),
                                    (bg_color, charcolor, SL_NONE, UL_NONE, W_NORMAL))

        def hide_numbering():
            for para in code_block:
                if p.match(para.String):
                    para.CharHeight = codecharheight
                    para.String = para.String[lenno:]

        if show:
            logger.debug("Showing code block numbering.")
            bg_color = self.to_int(char_bg_color) or -1
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
            lines = code_block.String.splitlines()
            if len(lines) == 1:  # this condition prevents to treat lines separated by carriage returns as a single line
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
        self.undomanager.enterUndoContext("All CH2 attributes removed.")
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
            self.undomanager.leaveUndoContext()
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
