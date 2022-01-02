# Code Highlighter
Code snippet highlighter for LibreOffice.

## INSTALLATION
Open LibreOffice, go to `Tools` -> `Extension Manager...` and add the extension `codehighlighter.oxt`

You can download the extension either from the official LibreOffice extensions page or from [releases](https://github.com/jmzambon/libreoffice-code-highlighter/releases).
If you have downloaded the `codehighlighter.oxt.zip` file from GitHub releases, extract it before adding to the LibreOffice.

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

Multiselection is also supported.

#### Screenshots
*Menu items*

![Menu](screenshots/code-highlighter-menu.png?raw=true "Menu")

*Dialog*

![Dialog](screenshots/code-highlighter-dialog.png?raw=true "Dialog")

*Result*

![Result](screenshots/code-highlighter-result.png?raw=true "Result")

## LICENSE
 - GPL v3
