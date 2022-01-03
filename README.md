# Code Highlighter 2
Code snippet highlighter for LibreOffice.

Code Highlighter 2 is a forked of [Code Highlighter](https://github.com/slgobinath/libreoffice-code-highlighter), originally created by slgobinath but no more maintained. 

Code Highlighter 2 is built upon [Pygments](https://pygments.org/) python library. 

## INSTALLATION

### Install Dependencies (Linux Users)
**Note: Close all the LibreOffice products before installing the dependencies**

Install libreoffice-script-provider-python
```
sudo apt-get install libreoffice-script-provider-python
```
The above command is for Ubuntu and its derivatives. Other Linux users, may not need this package.
If you encounter any problems after installing the extension, please check whether you have this or similar pacakge.

### Install Extension
Open LibreOffice, go to `Tools` -> `Extension Manager...` and add the extension `codehighlighter2.oxt`

You can download the extension either from the official LibreOffice extensions page or from [releases](https://github.com/jmzambon/libreoffice-code-highlighter/releases).
If you have downloaded the `codehighlighter2.oxt.zip` file from GitHub releases, extract it before adding to the LibreOffice.

## USAGE
- Open a LibreOffice document.
- Copy and paste any code snippet where you want it
  - **Writer**: either in a text frame (preferred option), in a text box, in a table cell or even as plain text
  - **Calc**: either in  cell or in a text box
  - **Impress** and **Draw**: in a text box
- Select the objetc or the text containing the code snippet.
- *Tools -> Highlight Code*
  - in the dialog: select the language and style
- Alternatively *Tools -> Highlight Code (previous settings)*
  - does not open a dialog, but applies previous settings (persistent also between restarts of LO)

#### Tips
- Multiselection is supported: you can select several code snippets at the same time, even from different languages (choose "automatic" in this last case).
- Not all language aliases that are valide names for Code Highlighter 2 appear in the option dialog: if you don't find a language, try anyway to type its name in the combobox (try for example with "R" or with "Pascal").  

#### Screenshots
*Menu items (Writer)*

![Menu](screenshots/code-highlighter-menu.png?raw=true "Menu")

*Dialog*

![Dialog](screenshots/code-highlighter-dialog.png?raw=true "Dialog")

*Result*

![Result](screenshots/code-highlighter-result.png?raw=true "Result")

## LICENSE
 - GPL v3
