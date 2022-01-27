[33mcommit 09687830e3eec0de8f97e3470d292d46736ca076[m[33m ([m[1;36mHEAD -> [m[1;32mmaster[m[33m)[m
Author: jmzambon <jeanmarczambon@gmail.com>
Date:   Mon Jan 24 22:45:11 2022 +0100

    Code enhancements.

[1mdiff --git a/codehighlighter/python/highlight.py b/codehighlighter/python/highlight.py[m
[1mindex f1934ff..21afb91 100644[m
[1m--- a/codehighlighter/python/highlight.py[m
[1m+++ b/codehighlighter/python/highlight.py[m
[36m@@ -126,7 +126,6 @@[m [mclass CodeHighlighter(unohelper.Base, XJobExecutor, XDialogEventHandler):[m
             self.frame = self.doc.CurrentController.Frame[m
             self.dispatcher = self.create("com.sun.star.frame.DispatchHelper")[m
             self.strings = ch2_i18n.getstrings(ctx)[m
[31m-            self.dialog = self.create_dialog()[m
             self.nolocale = Locale("zxx", "", "")[m
         except Exception:[m
             logging.exception("")[m
[36m@@ -258,17 +257,18 @@[m [mclass CodeHighlighter(unohelper.Base, XJobExecutor, XDialogEventHandler):[m
     def choose_options(self):[m
         # get options choice[m
         # 0: canceled, 1: OK[m
[31m-        # self.dialog.setVisible(True)[m
[31m-        if self.dialog.execute() == 0:[m
[32m+[m[32m        # dialog.setVisible(True)[m
[32m+[m[32m        dialog = self.create_dialog()[m
[32m+[m[32m        if dialog.execute() == 0:[m
             logging.debug("Dialog canceled.")[m
             return False[m
[31m-        lang = self.dialog.getControl('cb_lang').Text.strip() or 'automatic'[m
[31m-        style = self.dialog.getControl('cb_style').Text.strip() or 'default'[m
[31m-        colorize_bg = self.dialog.getControl('check_col_bg').State[m
[31m-        show_linenb = self.dialog.getControl('check_linenb').State[m
[31m-        nb_start = int(self.dialog.getControl('nb_start').Value)[m
[31m-        nb_ratio = int(self.dialog.getControl('nb_ratio').Value)[m
[31m-        nb_sep = self.dialog.getControl('nb_sep').Text[m
[32m+[m[32m        lang = dialog.getControl('cb_lang').Text.strip() or 'automatic'[m
[32m+[m[32m        style = dialog.getControl('cb_style').Text.strip() or 'default'[m
[32m+[m[32m        colorize_bg = dialog.getControl('check_col_bg').State[m
[32m+[m[32m        show_linenb = dialog.getControl('check_linenb').State[m
[32m+[m[32m        nb_start = int(dialog.getControl('nb_start').Value)[m
[32m+[m[32m        nb_ratio = int(dialog.getControl('nb_ratio').Value)[m
[32m+[m[32m        nb_sep = dialog.getControl('nb_sep').Text[m
 [m
         if lang != 'automatic' and lang.lower() not in self.all_lexer_aliases:[m
             self.msgbox(self.strings["errlang"])[m
[36m@@ -313,8 +313,9 @@[m [mclass CodeHighlighter(unohelper.Base, XJobExecutor, XDialogEventHandler):[m
         style = styles.get_style_by_name(stylename)[m
         bg_color = style.background_color if self.options['ColorizeBackground'] else None[m
 [m
[31m-        self.doc.lockControllers()[m
[31m-        logging.debug("Controllers locked.")[m
[32m+[m[32m        if not self.doc.hasControllersLocked():[m
[32m+[m[32m            self.doc.lockControllers()[m
[32m+[m[32m            logging.debug("Controllers locked.")[m
         undomanager = self.doc.UndoManager[m
         hascode = False[m
         try:[m
