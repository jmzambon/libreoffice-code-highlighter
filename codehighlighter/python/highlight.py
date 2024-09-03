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
import gettext
from com.sun.star.uno import RuntimeException
from com.sun.star.util import InvalidStateException


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
    logger.info(f"Pygments located in {pygments.__path__}.")
    logger.info(f"Lexers imported from {pygments.lexers.__file__}.")

    # uno
    import unohelper
    from com.sun.star.awt import Selection, XDialogEventHandler
    from com.sun.star.awt.FontWeight import NORMAL as W_NORMAL, BOLD as W_BOLD
    from com.sun.star.awt.FontSlant import NONE as SL_NONE, ITALIC as SL_ITALIC
    from com.sun.star.awt.FontUnderline import NONE as UL_NONE, SINGLE as UL_SINGLE
    from com.sun.star.awt.MessageBoxType import ERRORBOX, INFOBOX
    from com.sun.star.beans import PropertyValue
    from com.sun.star.container import ElementExistException
    from com.sun.star.document import XUndoAction
    from com.sun.star.drawing.FillStyle import SOLID as FS_SOLID  # , NONE as FS_NONE
    from com.sun.star.lang import Locale
    from com.sun.star.sheet.CellFlags import STRING as CF_STRING
    from com.sun.star.task import XJobExecutor
    from com.sun.star.xml import AttributeData

except Exception:
    logger.exception("Something went wrong while loading python modules:")   # see issue #28
    raise


CHARSTYLEID = "ch2_"
SNIPPETTAGID = CHARSTYLEID + "options"
INVALID_SELECTION = "Invalid"
SELECTED_PARASTYLE = {}


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
            self.parastyles = self.loadparastyles()
            self.cfg_access = self.create_cfg_access()
            self.options = self.load_options()
            self.extpath, self.extver = self.getextinfos()
            self.setlogger()
            logger.debug(f"Code Highlighter started from {self.doc.Title}.")
            logger.info(f"Loaded options = {self.options}.")
            self.frame = self.doc.CurrentController.Frame
            self.dispatcher = self.create("com.sun.star.frame.DispatchHelper")
            self.nolocale = Locale("zxx", "", "")
            self.inlinesnippet = False
            self.activepreviews = 0
            self.lexername = None

            # install gettext
            locdir = os.path.join(uno.fileUrlToSystemPath(self.extpath), "locales")
            logger.debug(f'Locales folder: {locdir}')
            # ps = self.create("com.sun.star.util.PathSubstitution")
            # vlang = ps.getSubstituteVariableValue("vlang")
            # lang = vlang.split("-")[0]
            # gtlang = gettext.translation('ch2', localedir=locdir, languages=[lang], fallback=True)
            gtlang = gettext.translation('ch2', localedir=locdir, fallback=True)
            gtlang.install(names=['_', 'ngettext'])
            _ = gtlang.gettext
            ngettext = gtlang.ngettext

        except Exception:
            logger.exception("Error initializing python class CodeHighlighter:")
            raise

    # XJobExecutor (https://www.openoffice.org/api/docs/common/ref/com/sun/star/task/XJobExecutor.html)
    def trigger(self, arg):
        logger.debug(f"Code Highlighter triggered with argument '{arg}'.")
        try:
            self.alert_on_empty_selection = True
            getattr(self, 'do_'+arg)()
        except Exception:
            logger.exception(f"Error triggering < self.do_{arg}() > function:")
            raise

    # XDialogEventHandler (http://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/XDialogEventHandler.html)
    def callHandlerMethod(self, dialog, event, method):
        logger.debug(f"Dialog handler action: '{method}'.")
        if method == "preview":
            focus = {1: 'cb_style', 2: 'nb_start'}
            self.do_preview(dialog)
            dialog.getControl(focus[dialog.Model.Step]).setFocus()
            return True
        elif method == "parastyle":
            try:
                lb_parastyle = dialog.getControl("lb_parastyle")
                locstylename = lb_parastyle.SelectedItem
                if locstylename == "":
                    self.msgbox(_("Please select a paragraph style."))
                else:
                    SELECTED_PARASTYLE[self.doc.RuntimeUID] = self.parastyles[locstylename]
                    dialog.endDialog(2)
                return True
            except Exception:
                traceback.print_exc()
        elif method in ("ev_linenb", "ev_charstyles"):
            self.expand_option(dialog, event.Selected, method)
            return True
        return False

    def getSupportedMethodNames(self):
        return 'preview', 'parastyle', 'ev_linenb', 'ev_charstyles'

    # main functions
    def do_highlight(self):
        '''Open option dialog and start code highlighting.'''

        self.selection = self.check_selection()
        if self.selection == INVALID_SELECTION:
            self.msgbox(_("Unsupported selection."))
        elif not self.selection and not self.parastyles:
            logger.debug("Current selection contains no text and no paragraph style is in use.")
            self.msgbox(_("Nothing to highlight."))
            return

        ret = self.choose_options()
        if ret == 2:
            self.inlinesnippet = False
            self.highlight_parastyle()
            return
        elif self.selection:
            logger.debug("Undoing existing previews on dialog closing.")
            while self.activepreviews:
                self.undomanager.undo()
                self.undomanager.clearRedo()
                self.activepreviews -= 1
            if ret == 1:
                logger.debug("Starting highlights.")
                for code_block in self.selection:
                    self.prepare_highlight(code_block)
        elif ret != 0:
            logger.debug("Current selection contains no text.")
            self.msgbox(_("Nothing to highlight."))

    def do_highlight_previous(self):
        '''Start code highlighting with current options as default.'''

        selection = self.check_selection()
        if selection == INVALID_SELECTION:
            self.msgbox(_("Unsupported selection."))
        elif selection:
            for code_block in selection:
                self.prepare_highlight(code_block)
        else:
            logger.debug("Current selection contains no text.")
            self.msgbox(_("Nothing to highlight."))

    def do_update(self):
        '''Update already highlighted snippets based on options stored in codeblock tags.
        Code-blocks must have been highlighted at least once with Code Highlighter 2.'''

        hasupdates = False
        selection = self.check_selection()
        if selection == INVALID_SELECTION:
            self.msgbox(_("Unsupported selection."))
        elif selection:
            for code_block in selection:
                ret = self.prepare_highlight(code_block, updatecode=True)
                if not hasupdates:
                    hasupdates = ret
            if not hasupdates:
                logger.debug("Selection is not updatable.")
                self.msgbox(_("Update impossible: no formatting attribute associated with this code."))
        else:
            logger.debug("Current selection contains no text.")
            self.msgbox(_("Nothing to highlight."))

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

    def loadparastyles(self):
        if self.doc.supportsService('com.sun.star.text.GenericTextDocument'):
            parastyles = self.doc.StyleFamilies.ParagraphStyles
            return {parastyles.getByName(name).DisplayName: name for name in parastyles.ElementNames if name != "Standard" and parastyles.getByName(name).isInUse()}
        else:
            return {}

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

        # set dialog strings
        controlnames = {"check_charstyles": (_("Use character ~styles"), _("When possible, code highlighting will be based on character styles.")),
                        "check_col_bg": (_("Set ~background from style"), _("Use the background color provided by the style.")),
                        "check_linenb": (_("Add ~line numbering"), _("Active or deactivate line numbers.")),
                        "lbl_nb_sep": (_("Separator"), _("Use \\t to insert tabulation")),
                        "lbl_nb_pad": (_("Padding symbol"), _("Character to fill the leading space (0 for 01 for example)")),
                        "lbl_cs_rootstyle": (_("Parent character style"), _("Use an existing character style as root style."))}
        for controlname in controlnames:
            dialog.getControl(controlname).Model.setPropertyValues(("Label", "HelpText"), controlnames[controlname])
        controlnames = {"label_lang": _("Language"), "label_style": _("Style"), "lbl_nb_start": _("Start at"), "lbl_nb_ratio": _("Height (%)"),
                        "label_parastyle": _("Highlight all codes formatted with paragraph style:"), "btn_parastyle": _("Highlight all"),
                        "pygments_ver": _("Build upon Pygments {}"), "preview": _("Preview")}
        for controlname in controlnames:
            dialog.getControl(controlname).Model.Label = controlnames[controlname]
        controlnames = {"nb_sep": _("Use \\t to insert tabulation"),
                        "nb_pad": _("Character to fill the leading space (0 for 01 for example)"),
                        "cs_rootstyle": _("Use an existing character style as root style."),
                        "lb_parastyle": _("Highlight every code snippet in the document that is formatted with the given paragraph style.")}
        for controlname in controlnames:
            dialog.getControl(controlname).Model.HelpText = controlnames[controlname]

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
        lb_parastyle = dialog.getControl('lb_parastyle')
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

        if self.parastyles:
            lb_parastyle.addItems(sorted(self.parastyles.keys(), key=str.casefold), 0)
            if SELECTED_PARASTYLE and SELECTED_PARASTYLE.get(self.doc.RuntimeUID, None) in self.parastyles:
                lb_parastyle.selectItem(SELECTED_PARASTYLE.get(self.doc.RuntimeUID, None), True)
        else:
            lb_parastyle.setEnable(False)
            if self.doc.supportsService('com.sun.star.text.GenericTextDocument'):
                lb_parastyle.Model.HelpText += _(" (There is currently no style in use.)")
            else:
                lb_parastyle.Model.HelpText += _(" (Writer only.)")
            dialog.getControl("btn_parastyle").setEnable(False)
            dialog.getControl("para_line").setEnable(False)

        check_col_bg.State = self.options['ColourizeBackground']
        check_charstyles.State = state1 = self.options['UseCharStyles']
        self.expand_option(dialog, state1, "ev_charstyles", False if state1==1 else True)
        check_linenb.State = state2 = self.options['ShowLineNumbers']
        self.expand_option(dialog, state2, "ev_linenb", False if state2==1 else True)
        nb_start.Value = self.options['LineNumberStart']
        nb_ratio.Value = self.options['LineNumberRatio']
        nb_sep.Text = self.options['LineNumberSeparator']
        nb_pad.Text = self.options['LineNumberPaddingSymbol']
        cs_rootstyle.Text = self.options['MasterCharStyle']
        logger.debug("--> filling controls ok.")

        dialog.Title = dialog.Title.format(self.extver)
        pygments_ver.Text = pygments_ver.Text.format(pygments.__version__)
        logger.debug("Dialog returned.")
        return dialog

    def getextinfos(self):
        pip = self.ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
        extensions = pip.getExtensionList()
        extid = "javahelps.codehighlighter"
        extpath = pip.getPackageLocation(extid)
        extver = ""
        for e in extensions:
            if extid in e:
                extver = e[1]
        return extpath, extver

    def get_options_from_dialog(self, dialog):
        opt = {}
        lang = dialog.getControl('cb_lang').Text.strip() or 'automatic'
        style = dialog.getControl('cb_style').Text.strip() or 'default'
        if lang != 'automatic' and lang.lower() not in self.all_lexer_aliases:
            self.msgbox(_("Unsupported language."))
        elif style not in self.all_styles:
            self.msgbox(_("Unknown style."))
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
        Dialog return values:
            0 = Canceled,
            1 = Highlight selection
            2 = Highlight paragraph style
        '''

        # dialog.setVisible(True)
        dialog = self.create_dialog()
        ret = dialog.execute()
        if ret == 0:
            logger.debug("Dialog canceled.")
            return ret

        choices = self.get_options_from_dialog(dialog)
        if not choices:
            return 0

        self.save_options(choices)
        logger.debug("Dialog validated and options saved.")
        logger.info(f"Updated options = {self.options}.")
        return ret

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
        self.lexername = lexer.name
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
            tok_style = style.style_for_token(ttype)
            if tok_style.get('italic', ''):
                newcharstyle.CharPosture = SL_ITALIC
            if tok_style.get('noitalic', ''):
                newcharstyle.CharPosture = SL_NONE
            if tok_style.get('bold', ''):
                newcharstyle.CharWeight = W_BOLD
            if tok_style.get('nobold', ''):
                newcharstyle.CharWeight = W_NORMAL
            if tok_style.get('underline', ''):
                newcharstyle.CharUnderline = UL_SINGLE
            if tok_style.get('nounderline', ''):
                newcharstyle.CharUnderline = UL_NONE
            if tok_style.get("bgcolor", ''):
                newcharstyle.CharBackColor = self.to_int(tok_style["bgcolor"])
            if tok_style.get("color", ''):
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
                    self.checkinlinesnippet(code_block)
                    if self.inlinesnippet:
                        udas = code_block.TextUserDefinedAttributes
                    else:
                        udas = code_block.ParaUserDefinedAttributes
                    if udas and SNIPPETTAGID in udas:
                        cursor = self.ensure_paragraphs(code_block)
                        self.doc.CurrentController.select(cursor)
                        code_blocks.append(self.doc.CurrentSelection[0])
                    elif code_block.TextFrame and code_block.Text.String.strip():
                        code_blocks.append(code_block.TextFrame)
                    elif code_block.TextTable and code_block.Text.String.strip():
                        cellname = code_block.Cell.CellName
                        cell = code_block.TextTable.getCellByName(cellname)
                        code_blocks.append(cell)

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
            elif code_block.ImplementationName in ("SwXTextRange", "SwXTextCursor"):
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
                        cursor = self.ensure_paragraphs(code_block)
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
                        cursor = self.ensure_paragraphs(code_block)
                        lexer = self.getlexer(cursor)
                        self.undomanager.enterUndoContext(f"code highlight (lang: {lexer.name}, style: {stylename})")
                        self.show_line_numbers(code_block, False, isplaintext=True)
                        cursor = self.ensure_paragraphs(code_block)  # in case numbering was removed, code_block has changed
                        controller.select(cursor)
                        cursor.CharLocale = self.nolocale
                        char_bg_color = None
                        if bg_color and not self.inlinesnippet:
                            # ParaBackColor does not work anymore, and new FillProperties isn't available from API
                            # see https://bugs.documentfoundation.org/show_bug.cgi?id=99125
                            # so let's use the dispatcher as workaround
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
                        try:
                            self.undomanager.leaveUndoContext()
                        except InvalidStateException:
                            pass

            # TEXT FRAME
            elif code_block.ImplementationName == "SwXTextFrame":
                logger.debug("Dealing with a text frame")
                code_block = code_block

                if updatecode:
                    # Frame's UserDefinedAttributes can't be reverted with undo manager -> using text cursor instead
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
                        self.tagcodeblock(code_block, lexer.name)
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
                        try:
                            controller.ViewCursor.collapseToEnd()
                        except RuntimeException:
                            pass
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
            self.msgbox(_("Unsupported selection."))
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
        if self.charstylesavailable and self.options["UseCharStyles"]:
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
        if self.lexername.startswith("LLVM") and ':' in sep:    # see issue https://github.com/jmzambon/libreoffice-code-highlighter/issues/27
            sep = '\t'
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

        def getregexstring():
            padsymbol = re.escape(pad)
            regexstring = fr"^[\s|{padsymbol}]*\d+\W?[^\S\n]*"
            if self.lexername.startswith("LLVM"):    # see issue https://github.com/jmzambon/libreoffice-code-highlighter/issues/27
                regexstring = fr"^[\s|{padsymbol}]*\d+(?![\d:])\W?[^\S\n]*"
            return regexstring

        if show:
            logger.debug("Showing code block numbering.")
            bg_color = self.to_int(char_bg_color) or -1
            show_numbering()
        else:
            # check for existing line numbering and its width
            regexstring = getregexstring()
            p = re.compile(regexstring, re.MULTILINE)
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
                return c
            c.collapseToStart()   # do not remove, needed for self.highlight_parastyle()
            c.gotoStartOfParagraph(False)
            c.gotoRange(selected_code.End, True)
            c.gotoEndOfParagraph(True)
        else:
            if self.inlinesnippet:
                udas = selected_code.TextUserDefinedAttributes
                if SNIPPETTAGID in udas:
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
                if SNIPPETTAGID in udas:
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
        return c

    # dev tools
    def expand_option(self, dialog, selected, method, move=True):
        controls = {'ev_charstyles': {'show': ('cs_rootstyle', 'lbl_cs_rootstyle'),
                                      'move': ('para_line', 'label_parastyle', 'lb_parastyle',
                                               'btn_parastyle', 'pygments_ver', 'pygments_logo'),
                                      'offset': {0: -14, 1: 14} },
                    'ev_linenb': {'show': ('nb_start', 'nb_ratio', 'nb_pad', 'nb_sep',
                                           'lbl_nb_start', 'lbl_nb_ratio', 'lbl_nb_pad', 'lbl_nb_sep'),
                                  'move': ('check_charstyles', 'cs_rootstyle', 'lbl_cs_rootstyle',
                                           'para_line', 'label_parastyle', 'lb_parastyle', 'btn_parastyle',
                                           'pygments_ver', 'pygments_logo'),
                                  'offset': {0: -72, 1: 72} }
                    }
        if method in controls:
            for controlname in controls[method]['show']:
                dialog.getControl(controlname).setVisible(selected)
            if move:
                offset = controls[method]['offset'][selected]
                dialog.Model.Height += offset
                for controlname in controls[method]['move']:
                    dialog.getControl(controlname).Model.PositionY += offset

    def highlight_parastyle(self):
        def finish_code_block(cursor):
            if cursor.Cell and cursor.String == cursor.Text.String:
                return cursor.Cell
            elif cursor.TextFrame and cursor.String == cursor.Text.String:
                return cursor.TextFrame
            else:
                return cursor

        def browse_all_paras(container=None):
            isroot = None
            if not container:
                isroot = True
                container = self.doc.Text

            cursor = None
            for para in container:
                if not para.supportsService('com.sun.star.text.Paragraph'):
                    continue
                if para.ParaStyleName == SELECTED_PARASTYLE[self.doc.RuntimeUID]:
                    if not cursor:
                        cursor = container.createTextCursorByRange(para.Start)
                    else:
                        cursor.gotoRange(para.End, True)
                elif cursor:

                    code_blocks.append(cursor)
                    cursor = None
            # if code_block is also the last paragraph of the text
            if cursor:
                code_blocks.append(cursor)
            # browse frames and text tables
            if isroot:
                for frame in self.doc.TextFrames:
                    browse_all_paras(frame)

                for table in self.doc.TextTables:
                    cellnames = table.CellNames
                    for cellname in cellnames:
                        cell = table.getCellByName(cellname)
                        browse_all_paras(cell)

        code_blocks = []
        browse_all_paras()
        if code_blocks:
            sel = self.doc.CurrentSelection
            for code_block in code_blocks:
                self.prepare_highlight(finish_code_block(code_block))
            message = ngettext("{} code snippet has been formatted.",
                               "{} code snippets have been formatted.",
                               len(code_blocks))
            self.doc.CurrentController.select(sel)
            self.msgbox(message.format(len(code_blocks)), boxtype=INFOBOX, title=_("Highlight all"))

    def update_all(self, usetags):
        '''Update all formatted code in the active document.
        DO NOT PUBLISH, ALPHA VERSION'''
        def highlight_snippet(code_block, udas):
            if usetags:
                options = literal_eval(udas.getByName(SNIPPETTAGID).Value)
                self.options.update(options)
            self.doc.CurrentController.select(code_block)
            logger.debug(f'Updating snippet (type: {code_block.ImplementationName})')
            self.prepare_highlight(code_block)

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
g_ImplementationHelper.addImplementation(
    CodeHighlighter, "ooo.ext.code-highlighter.impl", ("ooo.ext.code-highlighter",),)


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


def update_all_from_tags(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.update_all(True)


def update_all_from_options(event=None):
    ctx = XSCRIPTCONTEXT.getComponentContext()
    highlighter = CodeHighlighter(ctx)
    highlighter.update_all(False)
