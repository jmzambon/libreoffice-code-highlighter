"""
pygments.lexers._qlik_builtins
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Qlik builtins.

:copyright: Copyright 2006-2024 by the Pygments team, see AUTHORS.
:license: BSD, see LICENSE for details.
"""

# operators
#   see https://help.qlik.com/en-US/sense/August2021/Subsystems/Hub/Content/Sense_Hub/Scripting/Operators/operators.htm
OPERATORS_LIST = {
    "words": [
        # Bit operators
        "bitnot",
        "bitand",
        "bitor",
        "bitxor",
        # Logical operators
        "and",
        "or",
        "not",
        "xor",
        # Relational operators
        "precedes",
        "follows",
        # String operators
        "like",
    ],
    "symbols": [
        # Bit operators
        ">>",
        "<<",
        # Logical operators
        # Numeric operators
        "+",
        "-",
        "/",
        "*",
        # Relational operators
        "<",
        "<=",
        ">",
        ">=",
        "=",
        "<>",
        # String operators
        "&",
    ],
}

# SCRIPT STATEMENTS
#   see https://help.qlik.com/en-US/sense/August2021/Subsystems/Hub/Content/Sense_Hub/Scripting/
STATEMENT_LIST = [
    # control statements
    "for",
    "each",
    "in",
    "next",
    "do",
    "while",
    "until",
    "unless",
    "loop",
    "return",
    "switch",
    "case",
    "default",
    "if",
    "else",
    "endif",
    "then",
    "end",
    "exit",
    "script",
    "switch",
    # prefixes
    "Add",
    "Buffer",
    "Concatenate",
    "Crosstable",
    "First",
    "Generic",
    "Hierarchy",
    "HierarchyBelongsTo",
    "Inner",
    "IntervalMatch",
    "Join",
    "Keep",
    "Left",
    "Mapping",
    "Merge",
    "NoConcatenate",
    "Outer",
    "Partial reload",
    "Replace",
    "Right",
    "Sample",
    "Semantic",
    "Unless",
    "When",
    # regular statements
    "Alias",  # alias ... as ...
    "as",
    "AutoNumber",
    "Binary",
    "Comment field",  # comment fields ... using ...
    "Comment fields",  # comment field ... with ...
    "using",
    "with",
    "Comment table",  # comment table ... with ...
    "Comment tables",  # comment tables ... using ...
    "Connect",
    "ODBC",  # ODBC CONNECT TO ...
    "OLEBD",  # OLEDB CONNECT TO ...
    "CUSTOM",  # CUSTOM CONNECT TO ...
    "LIB",  # LIB CONNECT TO ...
    "Declare",
    "Derive",
    "From",
    "explicit",
    "implicit",
    "Direct Query",
    "dimension",
    "measure",
    "Directory",
    "Disconnect",
    "Drop field",
    "Drop fields",
    "Drop table",
    "Drop tables",
    "Execute",
    "FlushLog",
    "Force",
    "capitalization",
    "case upper",
    "case lower",
    "case mixed",
    "Load",
    "distinct",
    "from",
    "inline",
    "resident",
    "from_field",
    "autogenerate",
    "extension",
    "where",
    "group by",
    "order by",
    "asc",
    "desc",
    "Let",
    "Loosen Table",
    "Map",
    "NullAsNull",
    "NullAsValue",
    "Qualify",
    "Rem",
    "Rename field",
    "Rename fields",
    "Rename table",
    "Rename tables",
    "Search",
    "include",
    "exclude",
    "Section",
    "access",
    "application",
    "Select",
    "Set",
    "Sleep",
    "SQL",
    "SQLColumns",
    "SQLTables",
    "SQLTypes",
    "Star",
    "Store",
    "Tag",
    "Trace",
    "Unmap",
    "Unqualify",
    "Untag",
    # Qualifiers
    "total",
]

# Script functions
#    see https://help.qlik.com/en-US/sense/August2021/Subsystems/Hub/Content/Sense_Hub/Scripting/functions-in-scripts-chart-expressions.htm
SCRIPT_FUNCTIONS = [
    # Basic aggregation functions in the data load script
    "FirstSortedValue",
    "Max",
    "Min",
    "Mode",
    "Only",
    "Sum",
    # Counter aggregation functions in the data load script
    "Count",
    "MissingCount",
    "NullCount",
    "NumericCount",
    "TextCount",
    # Financial aggregation functions in the data load script
    "IRR",
    "XIRR",
    "NPV",
    "XNPV",
    # Statistical aggregation functions in the data load script
    "Avg",
    "Correl",
    "Fractile",
    "FractileExc",
    "Kurtosis",
    "LINEST_B" "LINEST_df",
    "LINEST_f",
    "LINEST_m",
    "LINEST_r2",
    "LINEST_seb",
    "LINEST_sem",
    "LINEST_sey",
    "LINEST_ssreg",
    "Linest_ssresid",
    "Median",
    "Skew",
    "Stdev",
    "Sterr",
    "STEYX",
    # Statistical test functions
    "Chi2Test_chi2",
    "Chi2Test_df",
    "Chi2Test_p",
    # Two independent samples t-tests
    "ttest_conf",
    "ttest_df",
    "ttest_dif",
    "ttest_lower",
    "ttest_sig",
    "ttest_sterr",
    "ttest_t",
    "ttest_upper",
    # Two independent weighted samples t-tests
    "ttestw_conf",
    "ttestw_df",
    "ttestw_dif",
    "ttestw_lower",
    "ttestw_sig",
    "ttestw_sterr",
    "ttestw_t",
    "ttestw_upper",
    # One sample t-tests
    "ttest1_conf",
    "ttest1_df",
    "ttest1_dif",
    "ttest1_lower",
    "ttest1_sig",
    "ttest1_sterr",
    "ttest1_t",
    "ttest1_upper",
    # One weighted sample t-tests
    "ttest1w_conf",
    "ttest1w_df",
    "ttest1w_dif",
    "ttest1w_lower",
    "ttest1w_sig",
    "ttest1w_sterr",
    "ttest1w_t",
    "ttest1w_upper",
    # One column format functions
    "ztest_conf",
    "ztest_dif",
    "ztest_sig",
    "ztest_sterr",
    "ztest_z",
    "ztest_lower",
    "ztest_upper",
    # Weighted two-column format functions
    "ztestw_conf",
    "ztestw_dif",
    "ztestw_lower",
    "ztestw_sig",
    "ztestw_sterr",
    "ztestw_upper",
    "ztestw_z",
    # String aggregation functions in the data load script
    "Concat",
    "FirstValue",
    "LastValue",
    "MaxString",
    "MinString",
    # Synthetic dimension functions
    "ValueList",
    "ValueLoop",
    # Color functions
    "ARGB",
    "HSL",
    "RGB",
    "Color",
    "Colormix1",
    "Colormix2",
    "SysColor",
    "ColorMapHue",
    "ColorMapJet",
    "black",
    "blue",
    "brown",
    "cyan",
    "darkgray",
    "green",
    "lightblue",
    "lightcyan",
    "lightgray",
    "lightgreen",
    "lightmagenta",
    "lightred",
    "magenta",
    "red",
    "white",
    "yellow",
    # Conditional functions
    "alt",
    "class",
    "coalesce",
    "if",
    "match",
    "mixmatch",
    "pick",
    "wildmatch",
    # Counter functions
    "autonumber",
    "autonumberhash128",
    "autonumberhash256",
    "IterNo",
    "RecNo",
    "RowNo",
    # Integer expressions of time
    "second",
    "minute",
    "hour",
    "day",
    "week",
    "month",
    "year",
    "weekyear",
    "weekday",
    # Timestamp functions
    "now",
    "today",
    "LocalTime",
    # Make functions
    "makedate",
    "makeweekdate",
    "maketime",
    # Other date functions
    "AddMonths",
    "AddYears",
    "yeartodate",
    # Timezone functions
    "timezone",
    "GMT",
    "UTC",
    "daylightsaving",
    "converttolocaltime",
    # Set time functions
    "setdateyear",
    "setdateyearmonth",
    # In... functions
    "inyear",
    "inyeartodate",
    "inquarter",
    "inquartertodate",
    "inmonth",
    "inmonthtodate",
    "inmonths",
    "inmonthstodate",
    "inweek",
    "inweektodate",
    "inlunarweek",
    "inlunarweektodate",
    "inday",
    "indaytotime",
    # Start ... end functions
    "yearstart",
    "yearend",
    "yearname",
    "quarterstart",
    "quarterend",
    "quartername",
    "monthstart",
    "monthend",
    "monthname",
    "monthsstart",
    "monthsend",
    "monthsname",
    "weekstart",
    "weekend",
    "weekname",
    "lunarweekstart",
    "lunarweekend",
    "lunarweekname",
    "daystart",
    "dayend",
    "dayname",
    # Day numbering functions
    "age",
    "networkdays",
    "firstworkdate",
    "lastworkdate",
    "daynumberofyear",
    "daynumberofquarter",
    # Exponential and logarithmic
    "exp",
    "log",
    "log10",
    "pow",
    "sqr",
    "sqrt",
    # Count functions
    "GetAlternativeCount",
    "GetExcludedCount",
    "GetNotSelectedCount",
    "GetPossibleCount",
    "GetSelectedCount",
    # Field and selection functions
    "GetCurrentSelections",
    "GetFieldSelections",
    "GetObjectDimension",
    "GetObjectField",
    "GetObjectMeasure",
    # File functions
    "Attribute",
    "ConnectString",
    "FileBaseName",
    "FileDir",
    "FileExtension",
    "FileName",
    "FilePath",
    "FileSize",
    "FileTime",
    "GetFolderPath",
    "QvdCreateTime",
    "QvdFieldName",
    "QvdNoOfFields",
    "QvdNoOfRecords",
    "QvdTableName",
    # Financial functions
    "FV",
    "nPer",
    "Pmt",
    "PV",
    "Rate",
    # Formatting functions
    "ApplyCodepage",
    "Date",
    "Dual",
    "Interval",
    "Money",
    "Num",
    "Time",
    "Timestamp",
    # General numeric functions
    "bitcount",
    "div",
    "fabs",
    "fact",
    "frac",
    "sign",
    # Combination and permutation functions
    "combin",
    "permut",
    # Modulo functions
    "fmod",
    "mod",
    # Parity functions
    "even",
    "odd",
    # Rounding functions
    "ceil",
    "floor",
    "round",
    # Geospatial functions
    "GeoAggrGeometry",
    "GeoBoundingBox",
    "GeoCountVertex",
    "GeoInvProjectGeometry",
    "GeoProjectGeometry",
    "GeoReduceGeometry",
    "GeoGetBoundingBox",
    "GeoGetPolygonCenter",
    "GeoMakePoint",
    "GeoProject",
    # Interpretation functions
    "Date#",
    "Interval#",
    "Money#",
    "Num#",
    "Text",
    "Time#",
    "Timestamp#",
    # Field functions
    "FieldIndex",
    "FieldValue",
    "FieldValueCount",
    # Inter-record functions in the data load script
    "Exists",
    "LookUp",
    "Peek",
    "Previous",
    # Logical functions
    "IsNum",
    "IsText",
    # Mapping functions
    "ApplyMap",
    "MapSubstring",
    # Mathematical functions
    "e",
    "false",
    "pi",
    "rand",
    "true",
    # NULL functions
    "EmptyIsNull",
    "IsNull",
    "Null",
    # Basic range functions
    "RangeMax",
    "RangeMaxString",
    "RangeMin",
    "RangeMinString",
    "RangeMode",
    "RangeOnly",
    "RangeSum",
    # Counter range functions
    "RangeCount",
    "RangeMissingCount",
    "RangeNullCount",
    "RangeNumericCount",
    "RangeTextCount",
    # Statistical range functions
    "RangeAvg",
    "RangeCorrel",
    "RangeFractile",
    "RangeKurtosis",
    "RangeSkew",
    "RangeStdev",
    # Financial range functions
    "RangeIRR",
    "RangeNPV",
    "RangeXIRR",
    "RangeXNPV",
    # Statistical distribution
    "CHIDIST",
    "CHIINV",
    "NORMDIST",
    "NORMINV",
    "TDIST",
    "TINV",
    "FDIST",
    "FINV",
    # String functions
    "Capitalize",
    "Chr",
    "Evaluate",
    "FindOneOf",
    "Hash128",
    "Hash160",
    "Hash256",
    "Index",
    "KeepChar",
    "Left",
    "Len",
    "LevenshteinDist",
    "Lower",
    "LTrim",
    "Mid",
    "Ord",
    "PurgeChar",
    "Repeat",
    "Replace",
    "Right",
    "RTrim",
    "SubField",
    "SubStringCount",
    "TextBetween",
    "Trim",
    "Upper",
    # System functions
    "Author",
    "ClientPlatform",
    "ComputerName",
    "DocumentName",
    "DocumentPath",
    "DocumentTitle",
    "EngineVersion",
    "GetCollationLocale",
    "GetObjectField",
    "GetRegistryString",
    "IsPartialReload",
    "OSUser",
    "ProductVersion",
    "ReloadTime",
    "StateName",
    # Table functions
    "FieldName",
    "FieldNumber",
    "NoOfFields",
    "NoOfRows",
    "NoOfTables",
    "TableName",
    "TableNumber",
]

# System variables and constants
# see https://help.qlik.com/en-US/sense/August2021/Subsystems/Hub/Content/Sense_Hub/Scripting/work-with-variables-in-data-load-editor.htm
CONSTANT_LIST = [
    # System Variables
    "floppy",
    "cd",
    "include",
    "must_include",
    "hideprefix",
    "hidesuffix",
    "qvpath",
    "qvroot",
    "QvWorkPath",
    "QvWorkRoot",
    "StripComments",
    "Verbatim",
    "OpenUrlTimeout",
    "WinPath",
    "WinRoot",
    "CollationLocale",
    "CreateSearchIndexOnReload",
    # value handling variables
    "NullDisplay",
    "NullInterpret",
    "NullValue",
    "OtherSymbol",
    # Currency formatting
    "MoneyDecimalSep",
    "MoneyFormat",
    "MoneyThousandSep",
    # Number formatting
    "DecimalSep",
    "ThousandSep",
    "NumericalAbbreviation",
    # Time formatting
    "DateFormat",
    "TimeFormat",
    "TimestampFormat",
    "MonthNames",
    "LongMonthNames",
    "DayNames",
    "LongDayNames",
    "FirstWeekDay",
    "BrokenWeeks",
    "ReferenceDay",
    "FirstMonthOfYear",
    # Error variables
    "errormode",
    "scripterror",
    "scripterrorcount",
    "scripterrorlist",
    # Other
    "null",
]
