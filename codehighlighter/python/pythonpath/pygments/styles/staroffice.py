from pygments.style import Style
from pygments.token import Keyword, Name, Comment, String, Error, \
     Number, Operator, Generic

class StarOfficeStyle(Style):
    default_style = ""
    styles = {
        Comment:                '#808080',   # Gray
        Keyword:                '#0000FF',   # Blue
        Name:                   '#008000',   # Green
        Name.Function:          '#008000',   # Green
        Name.Class:             '#008000',   # Green
        Number:                 '#FF0000',   # Red
        String:                 '#FF0000'    # Red
    }