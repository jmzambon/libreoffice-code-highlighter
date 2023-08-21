# Code Highlighter 2
Code snippet highlighter for LibreOffice.

Code Highlighter 2 is a forked of [Code Highlighter](https://github.com/slgobinath/libreoffice-code-highlighter), originally created by slgobinath but no more maintained. 

Code Highlighter 2 is built upon [Pygments](https://pygments.org/) python library. 

## INSTALLATION

> If you get an error message at installation time, try to uninstall any previous version of Code Highlighter, then restart LibreOffice and reinstall the extension.

### Install Dependencies (Linux Users)
**Note: Close all the LibreOffice products before installing the dependencies**

Install libreoffice-script-provider-python
```
sudo apt-get install libreoffice-script-provider-python
```
The above command is for Ubuntu and its derivatives. Other Linux users, may not need this package.
If you encounter any problems after installing the extension, please check whether you have this or similar package.

### Install Extension
Open LibreOffice, go to `Tools` -> `Extension Manager...` and add the extension `codehighlighter2.oxt`

You can download the extension either from the official LibreOffice extensions page or from the [Github repository](https://github.com/jmzambon/libreoffice-code-highlighter) (_codehighlighter2.oxt_ file).

## USAGE
- Open a LibreOffice document.
- Copy and paste any code snippet where you want it
  - **Writer**: either in a text frame (preferred option), in a text box, in a table cell or even as plain text
  - **Calc**: either in  cell or in a text box
  - **Impress** and **Draw**: in a text box
- Select the object or the text containing the code snippet.
- *Format -> Code Highlighter 2 -> Highlight Code*
  - in the dialog: select the language, the style and side parameters if needed
- or *Format -> Code Highlighter 2 -> Highlight Code (previous settings)*
  - does not open a dialog, but applies previous settings (persistent also between restarts of LO)
- or *Format -> Code Highlighter 2 -> Update selection*
  - updates an already highlighted snippet with the formatting informations stored with it

#### Features and Tips
- All languages (more then 500) and styles (more then 40) provided by Pygments are supported.
> **Tip 1** CodeHighlighter2 contains two styles that are not part of Pygments: libreoffice-classic and libreoffice-dark, that make use of LibreOffice IDE color schemes (classic mode and dark mode). Code Highlighter also adds a "LibreOffice Basic" language, which is not a Pygments lexer but a convenient shortcut to VB.net lexer, which is perfect for parsing LOBasic.

> **Tip 2** Not all language aliases that are valid names for Code Highlighter 2 appear in the option dialog: if you don't find a language, try anyway to type its name in the combobox (try for example with "R" or with "Pascal").
- Supports multiselection.
> **Tip** Choose "automatic" to highlight from different languages at the same time.
- Supports line numbering.
> **Tip 1** Click the "More..." button to access line numbers options or character styles options.

> **Tip 2** Uncheck line numbering option to remove unwanted line numbers, due for example to copy-pasted code.
- Supports all modules excepted Base.
> **Tip** For long snippet, CodeHighlighter2 works faster with text and text frame in Writer.
- Supports in-document preview [2.4.11].

#### General behaviour
- Highlighting is applied to the selected object, that can be plain text, text frame, text shape, text table cell or calc cell.
- Highlight properties are hard-formatted for each token, parsed according to the choosen Pygments lexer (aka coding language). If someone prefers instead to make use of character styles, select the corresponding option in the dialog box (from v2.3.0, Writer only).
- When cursor is inside a text shape or a Calc cell, highlighting is applied to the whole shape or cell even if only a part of the text is selected.
- When cursor is inside a text frame, a text table cell or an already highlighted plain text*, highlighting is only applied to the selected text if any, otherwise to the whole frame, cell or plain text snippet.
- When highlighting applies to the selected text, it formats the entire paragraphs, even if selection starts after the paragraph start or ends before paragraph end, unless the selection is an inline snippet.
- Choosing "Update selection", the program will update highlighted code keeping the already applied options. If nothing is selected and the cursor is inside an already highlighted block, the whole block will be updated*.

<sub>\* To allow code update, Code Highlighter 2 stores the formatting options in the document as [User Defined Attributes](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1xml_1_1UserDefinedAttributesSupplier.html#a7c8de9b61fff54bb35d4203618828f32). If you're not comfortable with that, you can disable it by setting the 'StoreOptionsWithSnippet' option to 0 in advanced options (Options-> Advanced-> Open Expert Configuration-> ooo.ext.code-highlighter.Registry).</sub>

## Screenshots
*Menu items (Writer)*

![Menu](screenshots/code-highlighter-menu.png?raw=true "Menu")

*Dialog*

![Dialog](screenshots/code-highlighter-dialog.png?raw=true "Dialog")

*Result*

![Result](screenshots/code-highlighter-result.png?raw=true "Result")

## LICENSE
 - GPL v3
