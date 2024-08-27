# Code Highlighter 2
Code snippet highlighter for LibreOffice Writer, Calc, Impress and Draw.

Code Highlighter 2 is a fork of [Code Highlighter](https://github.com/slgobinath/libreoffice-code-highlighter), originally created by [slgobinath](https://github.com/slgobinath) and no longer maintained.

Code Highlighter 2 is built upon [Pygments](https://pygments.org/) Python library.

## INSTALLATION

> If you get an error message at installation time, try to uninstall any previous version of Code Highlighter, then restart LibreOffice and reinstall the extension.

### Install Dependencies (Linux Users)
**Note: Close all LibreOffice applications before installing the dependencies**

Install libreoffice-script-provider-python
```
sudo apt-get install libreoffice-script-provider-python
```
The above command is for Ubuntu and its derivatives. Other Linux users, may not need this package.
If you encounter any problems after installing the extension, please check whether you have this or a similar package.

### Install Extension
Open LibreOffice, go to `Tools` → `Extension…` and add the extension `codehighlighter2.oxt`

You can download the extension either from the [official LibreOffice Extensions page](https://extensions.libreoffice.org/en/extensions/show/5814) or from the [GitHub repository](https://github.com/jmzambon/libreoffice-code-highlighter) (_codehighlighter2.oxt_ file).

## USAGE
- Open a LibreOffice Writer document, Calc spreadsheet, Impress presentation, or Draw drawing.
- Copy and paste any code snippet where you want it
  - **Writer**: either in a text frame (preferred option), in a text box, in a table cell or even as plain text
  - **Calc**: either in  cell or in a text box
  - **Impress** and **Draw**: in a text box
- Select the object or the text containing the code snippet.
- *Format → Code Highlighter 2 → Highlight Code*
  - in the dialog: select the language, the style and side parameters if needed
- or *Format → Code Highlighter 2 → Highlight Code (previous settings)*
  - does not open a dialog, but applies previous settings (persistent also between restarts of LO)
- or *Format → Code Highlighter 2 → Update selection*
  - updates an already highlighted snippet with the formatting informations stored with it

#### Alternatively (Writer only)
[On suggestion from kompilainenn]
- Format all your snippets with a dedicated paragraph style.
- Choose *Format → Code Highlighter 2 → Highlight Code*.
- Select the paragraph style and press *Highlight all*.

#### Features
- Supports all languages (more than 500) and styles (more than 40) provided by Pygments.
- Supports multiselection.
- Supports line numbering.
- Supports all modules excepted Base.
- Supports direct formatting or character styles.
- Allow to disable background color. 
- Allow in-document preview [2.4.11].

#### General behavior
- Highlighting is applied to the selected object, that can be plain text, text frame, text shape, text table cell or Calc spreadsheet cell.
- Highlight properties are hard-formatted for each token, parsed according to the choosen Pygments lexer (aka coding language). If someone prefers instead to make use of character styles, select the corresponding option in the dialog box (from v2.3.0, Writer only).
- When cursor is inside a text shape or a Calc cell, highlighting is applied to the whole shape or cell even if only a part of the text is selected.
- When cursor is inside a text frame, a text table cell or an already highlighted plain text*, highlighting is only applied to the selected text if any, otherwise to the whole frame, cell or plain text snippet.
- When highlighting applies to the selected text, it formats the entire paragraphs, even if selection starts after the paragraph start or ends before paragraph end, unless the selection is an inline snippet.
- Choosing “Update selection”, the program will update highlighted code keeping the already applied options. If nothing is selected and the cursor is inside an already highlighted block, the whole block will be updated*.

<sub>\* To allow code update, Code Highlighter 2 stores the formatting options in the document as [User Defined Attributes](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1xml_1_1UserDefinedAttributesSupplier.html#a7c8de9b61fff54bb35d4203618828f32). If you are not comfortable with that, you can disable it by setting the 'StoreOptionsWithSnippet' option to 0 in advanced options (Options → Advanced → Open Expert Configuration → ooo.ext.code-highlighter.Registry).</sub>

#### Tips
- CodeHighlighter2 contains two styles that are not part of Pygments: libreoffice-classic and libreoffice-dark, that make use of LibreOffice IDE color schemes (classic mode and dark mode). Code Highlighter also adds a “LibreOffice Basic” language, which is not a Pygments lexer but a convenient shortcut to VB.net lexer, which is perfect for parsing LOBasic.
- Not all language aliases that are valid names for Code Highlighter 2 appear in the option dialog: if you are unable to find a language, try anyway to type its name in the combobox (try for example with “R” or with “Pascal”).
- Choose “automatic” to highlight from different languages at the same time.
- Click the “More…” button to access line numbers options or character styles options.
- Uncheck line numbering option to remove unwanted line numbers, due for example to copy-pasted code.
- For long snippet, CodeHighlighter2 works faster with text and text frame in Writer.

## Screenshots
### Menu items (Writer)

![Menu](screenshots/code-highlighter-menu.png?raw=true "Menu")

### Dialog

![Dialog](screenshots/code-highlighter-dialog.png?raw=true "Dialog")

### Result

![Result](screenshots/code-highlighter-result.png?raw=true "Result")

## LOCALIZATION
- en (English, US)
- fr (French - Français)
- hu (Hungarian - magyar)
- it (Italian - Italiana)

## LICENSE
 - GPL v3
