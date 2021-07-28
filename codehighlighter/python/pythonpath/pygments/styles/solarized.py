"""
    pygments.styles.solarized
    ~~~~~~~~~~~~~~~~~~~~~~~~~

    The Solarized style, inspired by Schoonover.

    :copyright: Copyright 2012 by the Shoji KUMAGAI, see AUTHORS.
    :license: MIT, see LICENSE for details.
"""

from pygments.style import Style
from pygments.token import (
    Comment,
    Error,
    Generic,
    Keyword,
    Literal,
    Name,
    Number,
    Operator,
    Other,
    Punctuation,
    String,
    Text,
    Whitespace,
)


def rgb2hex(r, g, b):
    return f"#{''.join(f'{v:02x}' for v in (r, g, b))}"


BASE03 = rgb2hex(0, 43, 54)
BASE02 = rgb2hex(7, 54, 66)
BASE01 = rgb2hex(88, 110, 117)
BASE00 = rgb2hex(101, 123, 131)
BASE0 = rgb2hex(131, 148, 150)
BASE1 = rgb2hex(147, 161, 161)
BASE2 = rgb2hex(238, 232, 213)
BASE3 = rgb2hex(253, 246, 227)
YELLOW = rgb2hex(181, 137, 0)
ORANGE = rgb2hex(203, 75, 22)
RED = rgb2hex(220, 50, 47)
MAGENTA = rgb2hex(211, 54, 130)
VIOLET = rgb2hex(108, 113, 196)
BLUE = rgb2hex(38, 139, 210)
CYAN = rgb2hex(42, 161, 152)
GREEN = rgb2hex(133, 153, 0)


def _style(name, color):
    return f"{name} {color}"


def bold(color):
    return _style("bold", color)


def italic(color):
    return _style("italic", color)


class DarkStyle(Style):
    """Solarized Dark Style"""

    background_color = BASE03
    default_style = ""

    styles = {
        Text: BASE0,  # class: ''
        Whitespace: BASE03,  # class: 'w'
        Error: RED,  # class: 'err'
        Other: BASE0,  # class: 'x'
        Comment: italic(BASE01),  # class: 'c'
        Comment.Multiline: italic(BASE01),  # class: 'cm'
        Comment.Preproc: italic(BASE01),  # class: 'cp'
        Comment.Single: italic(BASE01),  # class: 'c1'
        Comment.Special: italic(BASE01),  # class: 'cs'
        Keyword: GREEN,  # class: 'k'
        Keyword.Constant: GREEN,  # class: 'kc'
        Keyword.Declaration: GREEN,  # class: 'kd'
        Keyword.Namespace: ORANGE,  # class: 'kn'
        Keyword.Pseudo: ORANGE,  # class: 'kp'
        Keyword.Reserved: GREEN,  # class: 'kr'
        Keyword.Type: GREEN,  # class: 'kt'
        Operator: BASE0,  # class: 'o'
        Operator.Word: GREEN,  # class: 'ow'
        Name: BASE1,  # class: 'n'
        Name.Attribute: BASE0,  # class: 'na'
        Name.Builtin: BLUE,  # class: 'nb'
        Name.Builtin.Pseudo: bold(BLUE),  # class: 'bp'
        Name.Class: BLUE,  # class: 'nc'
        Name.Constant: YELLOW,  # class: 'no'
        Name.Decorator: ORANGE,  # class: 'nd'
        Name.Entity: ORANGE,  # class: 'ni'
        Name.Exception: ORANGE,  # class: 'ne'
        Name.Function: BLUE,  # class: 'nf'
        Name.Property: BLUE,  # class: 'py'
        Name.Label: BASE0,  # class: 'nc'
        Name.Namespace: YELLOW,  # class: 'nn'
        Name.Other: BASE0,  # class: 'nx'
        Name.Tag: GREEN,  # class: 'nt'
        Name.Variable: ORANGE,  # class: 'nv'
        Name.Variable.Class: BLUE,  # class: 'vc'
        Name.Variable.Global: BLUE,  # class: 'vg'
        Name.Variable.Instance: BLUE,  # class: 'vi'
        Number: CYAN,  # class: 'm'
        Number.Float: CYAN,  # class: 'mf'
        Number.Hex: CYAN,  # class: 'mh'
        Number.Integer: CYAN,  # class: 'mi'
        Number.Integer.Long: CYAN,  # class: 'il'
        Number.Oct: CYAN,  # class: 'mo'
        Literal: BASE0,  # class: 'l'
        Literal.Date: BASE0,  # class: 'ld'
        Punctuation: BASE0,  # class: 'p'
        String: CYAN,  # class: 's'
        String.Backtick: CYAN,  # class: 'sb'
        String.Char: CYAN,  # class: 'sc'
        String.Doc: CYAN,  # class: 'sd'
        String.Double: CYAN,  # class: 's2'
        String.Escape: ORANGE,  # class: 'se'
        String.Heredoc: CYAN,  # class: 'sh'
        String.Interpol: ORANGE,  # class: 'si'
        String.Other: CYAN,  # class: 'sx'
        String.Regex: CYAN,  # class: 'sr'
        String.Single: CYAN,  # class: 's1'
        String.Symbol: CYAN,  # class: 'ss'
        Generic: BASE0,  # class: 'g'
        Generic.Deleted: BASE0,  # class: 'gd'
        Generic.Emph: BASE0,  # class: 'ge'
        Generic.Error: BASE0,  # class: 'gr'
        Generic.Heading: BASE0,  # class: 'gh'
        Generic.Inserted: BASE0,  # class: 'gi'
        Generic.Output: BASE0,  # class: 'go'
        Generic.Prompt: BASE0,  # class: 'gp'
        Generic.Strong: BASE0,  # class: 'gs'
        Generic.Subheading: BASE0,  # class: 'gu'
        Generic.Traceback: BASE0,  # class: 'gt'
    }


class LightStyle(Style):
    """Solarized Light style"""

    background_color = BASE3
    default_style = ""

    styles = {
        Text: BASE00,  # class: ''
        Whitespace: BASE3,  # class: 'w'
        Error: RED,  # class: 'err'
        Other: BASE00,  # class: 'x'
        Comment: italic(BASE1),  # class: 'c'
        Comment.Multiline: italic(BASE1),  # class: 'cm'
        Comment.Preproc: italic(BASE1),  # class: 'cp'
        Comment.Single: italic(BASE1),  # class: 'c1'
        Comment.Special: italic(BASE1),  # class: 'cs'
        Keyword: GREEN,  # class: 'k'
        Keyword.Constant: GREEN,  # class: 'kc'
        Keyword.Declaration: GREEN,  # class: 'kd'
        Keyword.Namespace: ORANGE,  # class: 'kn'
        Keyword.Pseudo: ORANGE,  # class: 'kp'
        Keyword.Reserved: GREEN,  # class: 'kr'
        Keyword.Type: GREEN,  # class: 'kt'
        Operator: BASE00,  # class: 'o'
        Operator.Word: GREEN,  # class: 'ow'
        Name: BASE01,  # class: 'n'
        Name.Attribute: BASE00,  # class: 'na'
        Name.Builtin: BLUE,  # class: 'nb'
        Name.Builtin.Pseudo: bold(BLUE),  # class: 'bp'
        Name.Class: BLUE,  # class: 'nc'
        Name.Constant: YELLOW,  # class: 'no'
        Name.Decorator: ORANGE,  # class: 'nd'
        Name.Entity: ORANGE,  # class: 'ni'
        Name.Exception: ORANGE,  # class: 'ne'
        Name.Function: BLUE,  # class: 'nf'
        Name.Property: BLUE,  # class: 'py'
        Name.Label: BASE00,  # class: 'nc'
        Name.Namespace: YELLOW,  # class: 'nn'
        Name.Other: BASE00,  # class: 'nx'
        Name.Tag: GREEN,  # class: 'nt'
        Name.Variable: ORANGE,  # class: 'nv'
        Name.Variable.Class: BLUE,  # class: 'vc'
        Name.Variable.Global: BLUE,  # class: 'vg'
        Name.Variable.Instance: BLUE,  # class: 'vi'
        Number: CYAN,  # class: 'm'
        Number.Float: CYAN,  # class: 'mf'
        Number.Hex: CYAN,  # class: 'mh'
        Number.Integer: CYAN,  # class: 'mi'
        Number.Integer.Long: CYAN,  # class: 'il'
        Number.Oct: CYAN,  # class: 'mo'
        Literal: BASE00,  # class: 'l'
        Literal.Date: BASE00,  # class: 'ld'
        Punctuation: BASE00,  # class: 'p'
        String: CYAN,  # class: 's'
        String.Backtick: CYAN,  # class: 'sb'
        String.Char: CYAN,  # class: 'sc'
        String.Doc: CYAN,  # class: 'sd'
        String.Double: CYAN,  # class: 's2'
        String.Escape: ORANGE,  # class: 'se'
        String.Heredoc: CYAN,  # class: 'sh'
        String.Interpol: ORANGE,  # class: 'si'
        String.Other: CYAN,  # class: 'sx'
        String.Regex: CYAN,  # class: 'sr'
        String.Single: CYAN,  # class: 's1'
        String.Symbol: CYAN,  # class: 'ss'
        Generic: BASE00,  # class: 'g'
        Generic.Deleted: BASE00,  # class: 'gd'
        Generic.Emph: BASE00,  # class: 'ge'
        Generic.Error: BASE00,  # class: 'gr'
        Generic.Heading: BASE00,  # class: 'gh'
        Generic.Inserted: BASE00,  # class: 'gi'
        Generic.Output: BASE00,  # class: 'go'
        Generic.Prompt: BASE00,  # class: 'gp'
        Generic.Strong: BASE00,  # class: 'gs'
        Generic.Subheading: BASE00,  # class: 'gu'
        Generic.Traceback: BASE00,  # class: 'gt'
    }
