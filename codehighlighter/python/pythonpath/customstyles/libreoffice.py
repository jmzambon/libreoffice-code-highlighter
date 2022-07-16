"""
    pygments.styles.libreoffice
    ~~~~~~~~~~~~~~~~~~~~

    LibreOffice basic IDE styles, classic and dark.
    Created for the LibreOffice plugin code-highlighter-2

    :copyright: public domain.
    :license: BSD, see LICENSE for details.
"""

from pygments.style import Style
from pygments.token import Comment, Error, Literal, Name, Token


class LibreOfficeStyle(Style):
    styles = {
        Token:                  '#000080',   # Blue
        Comment:                '#808080',   # Gray
        Error:                  '#800000',   # Lightred
        Literal:                '#ff0000',   # Red
        Name:                   '#008000',   # Green
    }


class LibreOfficeDarkStyle(Style):
    background_color = '#333333'
    styles = {
        Token:                  '#b4c7dc',
        Comment:                '#eeeeee',
        Error:                  '#ff3838',
        Literal:                '#ffa6a6',
        Name:                   '#dde8cb',
    }

