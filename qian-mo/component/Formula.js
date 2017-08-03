var Formula = function (sheet) {
    this.sheet = sheet
    //
// Formula constants for parsing:
//

    this.ParseState = {
        num: 1,
        alpha: 2,
        coord: 3,
        string: 4,
        stringquote: 5,
        numexp1: 6,
        numexp2: 7,
        alphanumeric: 8,
        specialvalue: 9
    };

    this.TokenType = {num: 1, coord: 2, op: 3, name: 4, error: 5, string: 6, space: 7};

    this.CharClass = {
        num: 1,
        numstart: 2,
        op: 3,
        eof: 4,
        alpha: 5,
        incoord: 6,
        error: 7,
        quote: 8,
        space: 9,
        specialstart: 10
    };

    this.CharClassTable = {
        " ": 9,
        "!": 3,
        '"': 8,
        "#": 10,
        "$": 6,
        "%": 3,
        "&": 3,
        "(": 3,
        ")": 3,
        "*": 3,
        "+": 3,
        ",": 3,
        "-": 3,
        ".": 2,
        "/": 3,
        "0": 1,
        "1": 1,
        "2": 1,
        "3": 1,
        "4": 1,
        "5": 1,
        "6": 1,
        "7": 1,
        "8": 1,
        "9": 1,
        ":": 3,
        "<": 3,
        "=": 3,
        ">": 3,
        "A": 5,
        "B": 5,
        "C": 5,
        "D": 5,
        "E": 5,
        "F": 5,
        "G": 5,
        "H": 5,
        "I": 5,
        "J": 5,
        "K": 5,
        "L": 5,
        "M": 5,
        "N": 5,
        "O": 5,
        "P": 5,
        "Q": 5,
        "R": 5,
        "S": 5,
        "T": 5,
        "U": 5,
        "V": 5,
        "W": 5,
        "X": 5,
        "Y": 5,
        "Z": 5,
        "^": 3,
        "_": 5,
        "a": 5,
        "b": 5,
        "c": 5,
        "d": 5,
        "e": 5,
        "f": 5,
        "g": 5,
        "h": 5,
        "i": 5,
        "j": 5,
        "k": 5,
        "l": 5,
        "m": 5,
        "n": 5,
        "o": 5,
        "p": 5,
        "q": 5,
        "r": 5,
        "s": 5,
        "t": 5,
        "u": 5,
        "v": 5,
        "w": 5,
        "x": 5,
        "y": 5,
        "z": 5
    };

    this.UpperCaseTable = {
        "a": "A",
        "b": "B",
        "c": "C",
        "d": "D",
        "e": "E",
        "f": "F",
        "g": "G",
        "h": "H",
        "i": "I",
        "j": "J",
        "k": "K",
        "l": "L",
        "m": "M",
        "n": "N",
        "o": "O",
        "p": "P",
        "q": "Q",
        "r": "R",
        "s": "S",
        "t": "T",
        "u": "U",
        "v": "V",
        "w": "W",
        "x": "X",
        "y": "Y",
        "z": "Z"
    }

    this.SpecialConstants = { // names that turn into constants for name lookup
        "#NULL!": "0,e#NULL!", "#NUM!": "0,e#NUM!", "#DIV/0!": "0,e#DIV/0!", "#VALUE!": "0,e#VALUE!",
        "#REF!": "0,e#REF!", "#NAME?": "0,e#NAME?"
    };


    // Operator Precedence table
    //
    // 1- !, 2- : ,, 3- M P, 4- %, 5- ^, 6- * /, 7- + -, 8- &, 9- < > = G(>=) L(<=) N(<>),
    // Negative value means Right Associative

    this.TokenPrecedence = {
        "!": 1,
        ":": 2, ",": 2,
        "M": -3, "P": -3,
        "%": 4,
        "^": 5,
        "*": 6, "/": 6,
        "+": 7, "-": 7,
        "&": 8,
        "<": 9, ">": 9, "G": 9, "L": 9, "N": 9
    };

    // Convert one-char token text to input text:

    this.TokenOpExpansion = {'G': '>=', 'L': '<=', 'M': '-', 'N': '<>', 'P': '+'};

    //
    // Information about the resulting value types when doing operations on values (used by LookupResultType)
    //
    // Each object entry is an object with specific types with result type info as follows:
    //
    //    'type1a': '|type2a:resulta|type2b:resultb|...
    //    Type of t* or n* matches any of those types not listed
    //    Results may be a type or the numbers 1 or 2 specifying to return type1 or type2


    this.TypeLookupTable = {
        unaryminus: {'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        unaryplus: {'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        unarypercent: {'n*': '|n:n%|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        plus: {
            'n%': '|n%:n%|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'nd': '|n%:n|nd:nd|nt:ndt|ndt:ndt|n$:n|n:nd|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'nt': '|n%:n|nd:ndt|nt:nt|ndt:ndt|n$:n|n:nt|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'ndt': '|n%:n|nd:ndt|nt:ndt|ndt:ndt|n$:n|n:ndt|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'n$': '|n%:n|nd:n|nt:n|ndt:n|n$:n$|n:n$|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'nl': '|n%:n|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'n': '|n%:n|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'b': '|n%:n%|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            't*': '|n*:e#VALUE!|t*:e#VALUE!|b:e#VALUE!|e*:2|',
            'e*': '|e*:1|n*:1|t*:1|b:1|'
        },
        concat: {
            't': '|t:t|th:th|tw:tw|tl:t|tr:tr|t*:2|e*:2|',
            'th': '|t:th|th:th|tw:t|tl:th|tr:t|t*:t|e*:2|',
            'tw': '|t:tw|th:t|tw:tw|tl:tw|tr:tw|t*:t|e*:2|',
            'tl': '|t:tl|th:th|tw:tw|tl:tl|tr:tr|t*:t|e*:2|',
            't*': '|t*:t|e*:2|',
            'e*': '|e*:1|n*:1|t*:1|'
        },
        oneargnumeric: {'n*': '|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        twoargnumeric: {
            'n*': '|n*:n|t*:e#VALUE!|e*:2|',
            'e*': '|e*:1|n*:1|t*:1|',
            't*': '|t*:e#VALUE!|n*:e#VALUE!|e*:2|'
        },
        propagateerror: {'n*': '|n*:2|e*:2|', 'e*': '|e*:2|', 't*': '|t*:2|e*:2|', 'b': '|b:2|e*:2|'}
    };
    this.scc = {
        s_parseerrexponent: "Improperly formed number exponent",
        s_parseerrchar: "Unexpected character in formula",
        s_parseerrstring: "Improperly formed string",
        s_parseerrspecialvalue: "Improperly formed special value",
        s_parseerrtwoops: "Error in formula (two operators inappropriately in a row)",
        s_parseerrmissingopenparen: "Missing open parenthesis in list with comma(s). ",
        s_parseerrcloseparennoopen: "Closing parenthesis without open parenthesis. ",
        s_parseerrmissingcloseparen: "Missing close parenthesis. ",
        s_parseerrmissingoperand: "Missing operand. ",
        s_parseerrerrorinformula: "Error in formula.",
        s_calcerrerrorvalueinformula: "Error value in formula",
        s_parseerrerrorinformulabadval: "Error in formula resulting in bad value",
        s_formularangeresult: "Formula results in range value:",
        s_calcerrnumericnan: "Formula results in an bad numeric value",
        s_calcerrnumericoverflow: "Numeric overflow",
        s_sheetunavailable: "Sheet unavailable:", // when FindSheetInCache returns null
        s_calcerrcellrefmissing: "Cell reference missing when expected.",
        s_calcerrsheetnamemissing: "Sheet name missing when expected.",
        s_circularnameref: "Circular name reference to name",
        s_calcerrunknownname: "Unknown name",
        s_calcerrincorrectargstofunction: "Incorrect arguments to function",
        s_sheetfuncunknownfunction: "Unknown function",
        s_sheetfunclnarg: "LN argument must be greater than 0",
        s_sheetfunclog10arg: "LOG10 argument must be greater than 0",
        s_sheetfunclogsecondarg: "LOG second argument must be numeric greater than 0",
        s_sheetfunclogfirstarg: "LOG first argument must be greater than 0",
        s_sheetfuncroundsecondarg: "ROUND second argument must be numeric",
        s_sheetfuncddblife: "DDB life must be greater than 1",
        s_sheetfuncslnlife: "SLN life must be greater than 1",
        s_InternalError: "Internal SocialCalc error (probably an internal bug): ", // hopefully unlikely, but a test failed
    }


    this.FunctionClassList = {
        //'statistics': [],
        'lookup': [],
        'datetime': [],
        'financial': [],
        'test': [],
        'math': [],
        'text': [],
        'stat': []
    }


    this.funcParem = {
        'CHOOSE': 'choose',
        'COLUMNS': 'range',
        'ROWS': 'range',
        'INDEX': 'index',
        'HLOOKUP': 'hlookup',
        'MATCH': 'match',
        'VLOOKUP': 'vlookup',
        'TODAY': '',
        'HOUR': 'v',
        'MINUTE': 'v',
        'SECOND': 'v',
        'DAY': 'v',
        'MONTH': 'v',
        'YEAR': 'v',
        'WEEKDAY': 'weekday',
        'TIME': 'hms',
        'DATE': 'date',
        'DDB': 'ddb',
        'SLN': 'csl',
        'SYD': 'cslp',
        'FV': 'fv',
        'NPER': 'nper',
        'PMT': 'pmt',
        'PV': 'pv',
        'RATE': 'rate',
        'NPV': 'npv',
        'IRR': 'irr',
        'AND': 'vn',
        'OR': 'vn',
        'NOT': 'v',
        'FALSE': '',
        'NA': '',
        'NOW': '',
        'PI': '',
        'TRUE': '',
        'T': 'v',
        'VALUE': 'v',
        'ISBLANK': 'v',
        'ISERR': 'v',
        'ISERROR': 'v',
        'ISLOGICAL': 'v',
        'ISNA': 'v',
        'ISNONTEXT': 'v',
        'ISNUMBER': 'v',
        'ISTEXT': 'v',
        'IF': 'iffunc',
        'ABS': 'v',
        'ACOS': 'v',
        'ASIN': 'v',
        'ATAN': 'v',
        'COS': 'v',
        'DEGREES': 'v',
        'EVEN': 'v',
        'EXP': 'v',
        'FACT': 'v',
        'INT': 'v',
        'LN': 'v',
        'LOG10': 'v',
        'ODD': 'v',
        'RADIANS': 'v',
        'SIN': 'v',
        'SQRT': 'v',
        'TAN': 'v',
        'ATAN2': 'xy',
        'MOD': '',
        'POWER': '',
        'TRUNC': 'valpre',
        'LOG': 'log',
        'ROUND': 'vp',
        'N': 'v',
        'FIND': 'find',
        'LEFT': 'tc',
        'LEN': 'txt',
        'LOWER': 'txt',
        'MID': 'mid',
        'PROPER': 'v',
        'REPLACE': 'replace',
        'REPT': 'tc',
        'RIGHT': 'tc',
        'SUBSTITUTE': 'subs',
        'TRIM': 'v',
        'UPPER': 'v',
        'EXACT': '',
        'COUNTIF': 'rangec',
        'SUMIF': 'sumif',
        'DAVERAGE': 'dfunc',
        'DCOUNT': 'dfunc',
        'DCOUNTA': 'dfunc',
        'DGET': 'dfunc',
        'DMAX': 'dfunc',
        'DMIN': 'dfunc',
        'DPRODUCT': 'dfunc',
        'DSTDEV': 'dfunc',
        'DSTDEVP': 'dfunc',
        'DSUM': 'dfunc',
        'DVAR': 'dfunc',
        'DVARP': 'dfunc',
        'AVERAGE': 'vn',
        'COUNT': 'vn',
        'COUNTA': 'vn',
        'COUNTBLANK': 'vn',
        'MAX': 'vn',
        'MIN': 'vn',
        'PRODUCT': 'vn',
        'STDEV': 'vn',
        'STDEVP': 'vn',
        'SUM': 'vn',
        'VAR': 'vn',
        'VARP': 'vn',


    }

    this.funcInfo = {
        s_fdef_ABS: 'Absolute value function. ',
        s_fdef_ACOS: 'Trigonometric arccosine function. ',
        s_fdef_AND: 'True if all arguments are true. ',
        s_fdef_ASIN: 'Trigonometric arcsine function. ',
        s_fdef_ATAN: 'Trigonometric arctan function. ',
        s_fdef_ATAN2: 'Trigonometric arc tangent function (result is in radians). ',
        s_fdef_AVERAGE: 'Averages the values. ',
        s_fdef_CHOOSE: 'Returns the value specified by the index. The values may be ranges of cells. ',
        s_fdef_COLUMNS: 'Returns the number of columns in the range. ',
        s_fdef_COS: 'Trigonometric cosine function (value is in radians). ',
        s_fdef_COUNT: 'Counts the number of numeric values, not blank, text, or error. ',
        s_fdef_COUNTA: 'Counts the number of non-blank values. ',
        s_fdef_COUNTBLANK: 'Counts the number of blank values. (Note: "" is not blank.) ',
        s_fdef_COUNTIF: 'Counts the number of number of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). ',
        s_fdef_DATE: 'Returns the appropriate date value given numbers for year, month, and day. For example: DATE(2006,2,1) for February 1, 2006. Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
        s_fdef_DAVERAGE: 'Averages the values in the specified field in records that meet the criteria. ',
        s_fdef_DAY: 'Returns the day of month for a date value. ',
        s_fdef_DCOUNT: 'Counts the number of numeric values, not blank, text, or error, in the specified field in records that meet the criteria. ',
        s_fdef_DCOUNTA: 'Counts the number of non-blank values in the specified field in records that meet the criteria. ',
        s_fdef_DDB: 'Returns the amount of depreciation at the given period of time (the default factor is 2 for double-declining balance).   ',
        s_fdef_DEGREES: 'Converts value in radians into degrees. ',
        s_fdef_DGET: 'Returns the value of the specified field in the single record that meets the criteria. ',
        s_fdef_DMAX: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DMIN: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DPRODUCT: 'Returns the result of multiplying the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DSTDEV: 'Returns the sample standard deviation of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DSTDEVP: 'Returns the standard deviation of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DSUM: 'Returns the sum of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DVAR: 'Returns the sample variance of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DVARP: 'Returns the variance of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_EVEN: 'Rounds the value up in magnitude to the nearest even integer. ',
        s_fdef_EXACT: 'Returns "true" if the values are exactly the same, including case, type, etc. ',
        s_fdef_EXP: 'Returns e raised to the value power. ',
        s_fdef_FACT: 'Returns factorial of the value. ',
        s_fdef_FALSE: 'Returns the logical value "false". ',
        s_fdef_FIND: 'Returns the starting position within string2 of the first occurrence of string1 at or after "start". If start is omitted, 1 is assumed. ',
        s_fdef_FV: 'Returns the future value of repeated payments of money invested at the given rate for the specified number of periods, with optional present value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
        s_fdef_HLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the row offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. ',
        s_fdef_HOUR: 'Returns the hour portion of a time or date/time value. ',
        s_fdef_IF: 'Results in true-value if logical-expression is TRUE or non-zero, otherwise results in false-value. ',
        s_fdef_INDEX: 'Returns a cell or range reference for the specified row and column in the range. If range is 1-dimensional, then only one of rownum or colnum are needed. If range is 2-dimensional and rownum or colnum are zero, a reference to the range of just the specified column or row is returned. You can use the returned reference value in a range, e.g., sum(A1:INDEX(A2:A10,4)). ',
        s_fdef_INT: 'Returns the value rounded down to the nearest integer (towards -infinity). ',
        s_fdef_IRR: 'Returns the interest rate at which the cash flows in the range have a net present value of zero. Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
        s_fdef_ISBLANK: 'Returns "true" if the value is a reference to a blank cell. ',
        s_fdef_ISERR: 'Returns "true" if the value is of type "Error" but not "NA". ',
        s_fdef_ISERROR: 'Returns "true" if the value is of type "Error". ',
        s_fdef_ISLOGICAL: 'Returns "true" if the value is of type "Logical" (true/false). ',
        s_fdef_ISNA: 'Returns "true" if the value is the error type "NA". ',
        s_fdef_ISNONTEXT: 'Returns "true" if the value is not of type "Text". ',
        s_fdef_ISNUMBER: 'Returns "true" if the value is of type "Number" (including logical values). ',
        s_fdef_ISTEXT: 'Returns "true" if the value is of type "Text". ',
        s_fdef_LEFT: 'Returns the specified number of characters from the text value. If count is omitted, 1 is assumed. ',
        s_fdef_LEN: 'Returns the number of characters in the text value. ',
        s_fdef_LN: 'Returns the natural logarithm of the value. ',
        s_fdef_LOG: 'Returns the logarithm of the value using the specified base. ',
        s_fdef_LOG10: 'Returns the base 10 logarithm of the value. ',
        s_fdef_LOWER: 'Returns the text value with all uppercase characters converted to lowercase. ',
        s_fdef_MATCH: 'Look for the matching value for the given value in the range and return position (the first is 1) in that range. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. If rangelookup is -1, act like 1 but the bracket is match>=value. ',
        s_fdef_MAX: 'Returns the maximum of the numeric values. ',
        s_fdef_MID: 'Returns the specified number of characters from the text value starting from the specified position. ',
        s_fdef_MIN: 'Returns the minimum of the numeric values. ',
        s_fdef_MINUTE: 'Returns the minute portion of a time or date/time value. ',
        s_fdef_MOD: 'Returns the remainder of the first value divided by the second. ',
        s_fdef_MONTH: 'Returns the month part of a date value. ',
        s_fdef_N: 'Returns the value if it is a numeric value otherwise an error. ',
        s_fdef_NA: 'Returns the #N/A error value which propagates through most operations. ',
        s_fdef_NOT: 'Returns FALSE if value is true, and TRUE if it is false. ',
        s_fdef_NOW: 'Returns the current date/time. ',
        s_fdef_NPER: 'Returns the number of periods at which payments invested each period at the given rate with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period) has the given present value. ',
        s_fdef_NPV: 'Returns the net present value of cash flows (which may be individual values and/or ranges) at the given rate. The flows are positive if income, negative if paid out, and are assumed at the end of each period. ',
        s_fdef_ODD: 'Rounds the value up in magnitude to the nearest odd integer. ',
        s_fdef_OR: 'True if any argument is true ',
        s_fdef_PI: 'The value 3.1415926... ',
        s_fdef_PMT: 'Returns the amount of each payment that must be invested at the given rate for the specified number of periods to have the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
        s_fdef_POWER: 'Returns the first value raised to the second value power. ',
        s_fdef_PRODUCT: 'Returns the result of multiplying the numeric values. ',
        s_fdef_PROPER: 'Returns the text value with the first letter of each word converted to uppercase and the others to lowercase. ',
        s_fdef_PV: 'Returns the present value of the given number of payments each invested at the given rate, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
        s_fdef_RADIANS: 'Converts value in degrees into radians. ',
        s_fdef_RATE: 'Returns the rate at which the given number of payments each invested at the given rate has the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
        s_fdef_REPLACE: 'Returns text1 with the specified number of characters starting from the specified position replaced by text2. ',
        s_fdef_REPT: 'Returns the text repeated the specified number of times. ',
        s_fdef_RIGHT: 'Returns the specified number of characters from the text value starting from the end. If count is omitted, 1 is assumed. ',
        s_fdef_ROUND: 'Rounds the value to the specified number of decimal places. If precision is negative, then round to powers of 10. The default precision is 0 (round to integer). ',
        s_fdef_ROWS: 'Returns the number of rows in the range. ',
        s_fdef_SECOND: 'Returns the second portion of a time or date/time value (truncated to an integer). ',
        s_fdef_SIN: 'Trigonometric sine function (value is in radians) ',
        s_fdef_SLN: 'Returns the amount of depreciation at each period of time using the straight-line method. ',
        s_fdef_SQRT: 'Square root of the value ',
        s_fdef_STDEV: 'Returns the sample standard deviation of the numeric values. ',
        s_fdef_STDEVP: 'Returns the standard deviation of the numeric values. ',
        s_fdef_SUBSTITUTE: 'Returns text1 with the all occurrences of oldtext replaced by newtext. If "occurrence" is present, then only that occurrence is replaced. ',
        s_fdef_SUM: 'Adds the numeric values. The values to the sum function may be ranges in the form similar to A1:B5. ',
        s_fdef_SUMIF: 'Sums the numeric values of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). If range2 is present, then range1 is tested and the corresponding range2 value is summed. ',
        s_fdef_SYD: 'Depreciation by Sum of Year\'s Digits method. ',
        s_fdef_T: 'Returns the text value or else a null string. ',
        s_fdef_TAN: 'Trigonometric tangent function (value is in radians) ',
        s_fdef_TIME: 'Returns the time value given the specified hour, minute, and second. ',
        s_fdef_TODAY: 'Returns the current date (an integer). Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
        s_fdef_TRIM: 'Returns the text value with leading, trailing, and repeated spaces removed. ',
        s_fdef_TRUE: 'Returns the logical value "true". ',
        s_fdef_TRUNC: 'Truncates the value to the specified number of decimal places. If precision is negative, truncate to powers of 10. ',
        s_fdef_UPPER: 'Returns the text value with all lowercase characters converted to uppercase. ',
        s_fdef_VALUE: 'Converts the specified text value into a numeric value. Various forms that look like numbers (including digits followed by %, forms that look like dates, etc.) are handled. This may not handle all of the forms accepted by other spreadsheets and may be locale dependent. ',
        s_fdef_VAR: 'Returns the sample variance of the numeric values. ',
        s_fdef_VARP: 'Returns the variance of the numeric values. ',
        s_fdef_VLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the column offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match>=value) instead of exact match. ',
        s_fdef_WEEKDAY: 'Returns the day of week specified by the date value. If type is 1 (the default), Sunday is day and Saturday is day 7. If type is 2, Monday is day 1 and Sunday is day 7. If type is 3, Monday is day 0 and Sunday is day 6. ',
        s_fdef_YEAR: 'Returns the year part of a date value. ',

        s_farg_v: "value",
        s_farg_vn: "value1, value2, ...",
        s_farg_xy: "valueX, valueY",
        s_farg_choose: "index, value1, value2, ...",
        s_farg_range: "range",
        s_farg_rangec: "range, criteria",
        s_farg_date: "year, month, day",
        s_farg_dfunc: "databaserange, fieldname, criteriarange",
        s_farg_ddb: "cost, salvage, lifetime, period [, factor]",
        s_farg_find: "string1, string2 [, start]",
        s_farg_fv: "rate, n, payment, [pv, [paytype]]",
        s_farg_hlookup: "value, range, row, [rangelookup]",
        s_farg_iffunc: "logical-expression, true-value, false-value",
        s_farg_index: "range, rownum, colnum",
        s_farg_irr: "range, [guess]",
        s_farg_tc: "text, count",
        s_farg_log: "value, base",
        s_farg_match: "value, range, [rangelookup]",
        s_farg_mid: "text, start, length",
        s_farg_nper: "rate, payment, pv, [fv, [paytype]]",
        s_farg_npv: "rate, value1, value2, ...",
        s_farg_pmt: "rate, n, pv, [fv, [paytype]]",
        s_farg_pv: "rate, n, payment, [fv, [paytype]]",
        s_farg_rate: "n, payment, pv, [fv, [paytype, [guess]]]",
        s_farg_replace: "text1, start, length, text2",
        s_farg_vp: "value, [precision]",
        s_farg_valpre: "value, precision",
        s_farg_csl: "cost, salvage, lifetime",
        s_farg_cslp: "cost, salvage, lifetime, period",
        s_farg_subs: "text1, oldtext, newtext [, occurrence]",
        s_farg_sumif: "range1, criteria [, range2]",
        s_farg_hms: "hour, minute, second",
        s_farg_txt: "text",
        s_farg_vlookup: "value, range, col, [rangelookup]",
        s_farg_weekday: "date, [type]",
        s_farg_dt: "date",
        s_farg_: ''
    }


/////////////////////////
//
// FUNCTION DEFINITIONS
//
// The standard function definitions follow.
//
// Note that some need SocialCalc.DetermineValueType to be defined.
//

    /*
    #
    # AVERAGE(v1,c1:c2,...)
    # COUNT(v1,c1:c2,...)
    # COUNTA(v1,c1:c2,...)
    # COUNTBLANK(v1,c1:c2,...)
    # MAX(v1,c1:c2,...)
    # MIN(v1,c1:c2,...)
    # PRODUCT(v1,c1:c2,...)
    # STDEV(v1,c1:c2,...)
    # STDEVP(v1,c1:c2,...)
    # SUM(v1,c1:c2,...)
    # VAR(v1,c1:c2,...)
    # VARP(v1,c1:c2,...)
    #
    # Calculate all of these and then return the desired one (overhead is in accessing not calculating)
    # If this routine is changed, check the dseries_functions, too.
    #
    */

    //var SocialCalc = {}
    //this.SocialCalcFormula = {}
    var SocialCalcFormula = this
    SocialCalcFormula.PushOperand = function (operand, t, v) {

        operand.push({type: t, value: v});

    }
    var FunctionClassList = this.FunctionClassList

    SocialCalcFormula.SeriesFunctions = function (fname, operand, foperand) {

        var value1, t, v1;

        var scf = SocialCalcFormula;
        //var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        var sum = 0;
        var resulttypesum = "";
        var count = 0;
        var counta = 0;
        var countblank = 0;
        var product = 1;
        var maxval;
        var minval;
        var mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                              // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

        while (foperand.length > 0) {
            value1 = scf.OperandValueAndType(foperand);
            t = value1.type.charAt(0);
            if (t == "n") count += 1;
            if (t != "b") counta += 1;
            if (t == "b") countblank += 1;

            if (t == "n") {
                v1 = value1.value - 0; // get it as a number
                sum += v1;
                product *= v1;
                maxval = (maxval != undefined) ? (v1 > maxval ? v1 : maxval) : v1;
                minval = (minval != undefined) ? (v1 < minval ? v1 : minval) : v1;
                if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
                    mk1 = v1;
                    sk1 = 0;
                }
                else { // Accumulate S sub 1 through n as per Knuth noted above
                    mk = mk1 + (v1 - mk1) / count;
                    sk = sk1 + (v1 - mk1) * (v1 - mk);
                    sk1 = sk;
                    mk1 = mk;
                }
                resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
            }
            else if (t == "e" && resulttypesum.charAt(0) != "e") {
                resulttypesum = value1.type;
            }
        }

        resulttypesum = resulttypesum || "n";

        switch (fname) {
            case "SUM":
                PushOperand(resulttypesum, sum);
                break;

            case "PRODUCT": // may handle cases with text differently than some other spreadsheets
                PushOperand(resulttypesum, product);
                break;

            case "MIN":
                PushOperand(resulttypesum, minval || 0);
                break;

            case "MAX":
                PushOperand(resulttypesum, maxval || 0);
                break;

            case "COUNT":
                PushOperand("n", count);
                break;

            case "COUNTA":
                PushOperand("n", counta);
                break;

            case "COUNTBLANK":
                PushOperand("n", countblank);
                break;

            case "AVERAGE":
                if (count > 0) {
                    PushOperand(resulttypesum, sum / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "STDEV":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "STDEVP":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / count));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "VAR":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / (count - 1));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "VARP":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;
        }

        return null;

    }

// Add to function list
    FunctionClassList['stat']["AVERAGE"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["COUNT"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["COUNTA"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["COUNTBLANK"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["MAX"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["MIN"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["PRODUCT"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["STDEV"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["STDEVP"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["SUM"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["VAR"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["VARP"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];

    /*
    #
    # DAVERAGE(databaserange, fieldname, criteriarange)
    # DCOUNT(databaserange, fieldname, criteriarange)
    # DCOUNTA(databaserange, fieldname, criteriarange)
    # DGET(databaserange, fieldname, criteriarange)
    # DMAX(databaserange, fieldname, criteriarange)
    # DMIN(databaserange, fieldname, criteriarange)
    # DPRODUCT(databaserange, fieldname, criteriarange)
    # DSTDEV(databaserange, fieldname, criteriarange)
    # DSTDEVP(databaserange, fieldname, criteriarange)
    # DSUM(databaserange, fieldname, criteriarange)
    # DVAR(databaserange, fieldname, criteriarange)
    # DVARP(databaserange, fieldname, criteriarange)
    #
    # Calculate all of these and then return the desired one (overhead is in accessing not calculating)
    # If this routine is changed, check the series_functions, too.
    #
    */

    SocialCalcFormula.DSeriesFunctions = function (fname, operand, foperand) {

        var value1, tostype, cr, dbrange, fieldname, criteriarange, dbinfo, criteriainfo;
        var fieldasnum, targetcol, i, j, k, cell, criteriafieldnums;
        var testok, criteriacr, criteria, testcol, testcr;
        var t;

        var scf = SocialCalcFormula;
        //var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        var value1 = {};

        var sum = 0;
        var resulttypesum = "";
        var count = 0;
        var counta = 0;
        var countblank = 0;
        var product = 1;
        var maxval;
        var minval;
        var mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                              // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

        dbrange = scf.TopOfStackValueAndType(foperand); // get a range
        fieldname = scf.OperandValueAndType(foperand); // get a value
        criteriarange = scf.TopOfStackValueAndType(foperand); // get a range

        if (dbrange.type != "range" || criteriarange.type != "range") {
            return scf.FunctionArgsError(fname, operand);
        }

        dbinfo = scf.DecodeRangeParts(dbrange.value);
        criteriainfo = scf.DecodeRangeParts(criteriarange.value);

        fieldasnum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, fieldname.value, fieldname.type);
        if (fieldasnum <= 0) {
            PushOperand("e#VALUE!", 0);
            return;
        }

        targetcol = dbinfo.col1num + fieldasnum - 1;
        criteriafieldnums = [];

        for (i = 0; i < criteriainfo.ncols; i++) { // get criteria field colnums
            cell = criteriainfo.sheetdata.GetAssuredCell(SocialCalc.crToCoord(criteriainfo.col1num + i, criteriainfo.row1num));
            criterianum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, cell.datavalue, cell.valuetype);
            if (criterianum <= 0) {
                PushOperand("e#VALUE!", 0);
                return;
            }
            criteriafieldnums.push(dbinfo.col1num + criterianum - 1);
        }

        for (i = 1; i < dbinfo.nrows; i++) { // go through each row of the database
            testok = false;
            CRITERIAROW:
                for (j = 1; j < criteriainfo.nrows; j++) { // go through each criteria row
                    for (k = 0; k < criteriainfo.ncols; k++) { // look at each column
                        criteriacr = SocialCalc.crToCoord(criteriainfo.col1num + k, criteriainfo.row1num + j); // where criteria is
                        cell = criteriainfo.sheetdata.GetAssuredCell(criteriacr);
                        criteria = cell.datavalue;
                        if (typeof criteria == "string" && criteria.length == 0) continue; // blank items are OK
                        testcol = criteriafieldnums[k];
                        testcr = SocialCalc.crToCoord(testcol, dbinfo.row1num + i); // cell to check
                        cell = criteriainfo.sheetdata.GetAssuredCell(testcr);
                        if (!scf.TestCriteria(cell.datavalue, cell.valuetype || "b", criteria)) {
                            continue CRITERIAROW; // does not meet criteria - check next row
                        }
                    }
                    testok = true; // met all the criteria
                    break CRITERIAROW;
                }
            if (!testok) {
                continue;
            }

            cr = SocialCalc.crToCoord(targetcol, dbinfo.row1num + i); // get cell of this row to do the function on
            cell = dbinfo.sheetdata.GetAssuredCell(cr);

            value1.value = cell.datavalue;
            value1.type = cell.valuetype;
            t = value1.type.charAt(0);
            if (t == "n") count += 1;
            if (t != "b") counta += 1;
            if (t == "b") countblank += 1;

            if (t == "n") {
                v1 = value1.value - 0; // get it as a number
                sum += v1;
                product *= v1;
                maxval = (maxval != undefined) ? (v1 > maxval ? v1 : maxval) : v1;
                minval = (minval != undefined) ? (v1 < minval ? v1 : minval) : v1;
                if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
                    mk1 = v1;
                    sk1 = 0;
                }
                else { // Accumulate S sub 1 through n as per Knuth noted above
                    mk = mk1 + (v1 - mk1) / count;
                    sk = sk1 + (v1 - mk1) * (v1 - mk);
                    sk1 = sk;
                    mk1 = mk;
                }
                resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
            }
            else if (t == "e" && resulttypesum.charAt(0) != "e") {
                resulttypesum = value1.type;
            }
        }

        resulttypesum = resulttypesum || "n";

        switch (fname) {
            case "DSUM":
                PushOperand(resulttypesum, sum);
                break;

            case "DPRODUCT": // may handle cases with text differently than some other spreadsheets
                PushOperand(resulttypesum, product);
                break;

            case "DMIN":
                PushOperand(resulttypesum, minval || 0);
                break;

            case "DMAX":
                PushOperand(resulttypesum, maxval || 0);
                break;

            case "DCOUNT":
                PushOperand("n", count);
                break;

            case "DCOUNTA":
                PushOperand("n", counta);
                break;

            case "DAVERAGE":
                if (count > 0) {
                    PushOperand(resulttypesum, sum / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DSTDEV":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DSTDEVP":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / count));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DVAR":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / (count - 1));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DVARP":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DGET":
                if (count == 1) {
                    PushOperand(resulttypesum, sum);
                }
                else if (count == 0) {
                    PushOperand("e#VALUE!", 0);
                }
                else {
                    PushOperand("e#NUM!", 0);
                }
                break;

        }

        return;

    }

    FunctionClassList['stat']["DAVERAGE"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DCOUNT"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DCOUNTA"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DGET"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DMAX"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DMIN"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DPRODUCT"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DSTDEV"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DSTDEVP"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DSUM"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DVAR"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DVARP"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];

    /*
    #
    # colnum = SocialCalcFormula.FieldToColnum(sheet, col1num, ncols, row1num, fieldname, fieldtype)
    #
    # If fieldname is a number, uses it, otherwise looks up string in cells in row to find field number
    #
    # If not found, returns 0.
    #
    */

    // SocialCalcFormula.FieldToColnum = function (col1num, ncols, row1num, fieldname, fieldtype) {
    //
    //     var colnum, cell, value;
    //
    //     if (fieldtype.charAt(0) == "n") { // number - return it if legal
    //         colnum = fieldname - 0; // make sure a number
    //         if (colnum <= 0 || colnum > ncols) {
    //             return 0;
    //         }
    //         return Math.floor(colnum);
    //     }
    //
    //     if (fieldtype.charAt(0) != "t") { // must be text otherwise
    //         return 0;
    //     }
    //
    //     fieldname = fieldname ? fieldname.toLowerCase() : "";
    //
    //     for (colnum = 0; colnum < ncols; colnum++) { // look through column headers for a match
    //         cell = sheet.GetAssuredCell(SocialCalc.crToCoord(col1num + colnum, row1num));
    //         value = cell.datavalue;
    //         value = (value + "").toLowerCase(); // ignore case
    //         if (value == fieldname) { // match
    //             return colnum + 1;
    //         }
    //     }
    //     return 0; // looked at all and no match
    //
    // }


    /*
    #
    # HLOOKUP(value, range, row, [rangelookup])
    # VLOOKUP(value, range, col, [rangelookup])
    # MATCH(value, range, [rangelookup])
    #
    */

    SocialCalcFormula.LookupFunctions = function (fname, operand, foperand) {

        var lookupvalue, range, offset, rangelookup, offsetvalue, rangeinfo;
        var c, r, cincr, rincr, previousOK, csave, rsave, cell, value, valuetype, cr, lookupvalue;

        var scf = SocialCalcFormula;
        //var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        lookupvalue = scf.OperandValueAndType(foperand);
        if (typeof lookupvalue.value == "string") {
            lookupvalue.value = lookupvalue.value.toLowerCase();
        }

        range = scf.TopOfStackValueAndType(foperand);

        rangelookup = 1; // default to true or 1
        if (fname == "MATCH") {
            if (foperand.length) {
                rangelookup = scf.OperandAsNumber(foperand);
                if (rangelookup.type.charAt(0) != "n") {
                    PushOperand("e#VALUE!", 0);
                    return;
                }
                if (foperand.length) {
                    scf.FunctionArgsError(fname, operand);
                    return 0;
                }
                rangelookup = rangelookup.value - 0;
            }
        }
        else {
            offsetvalue = scf.OperandAsNumber(foperand);
            if (offsetvalue.type.charAt(0) != "n") {
                PushOperand("e#VALUE!", 0);
                return;
            }
            offsetvalue = Math.floor(offsetvalue.value);
            if (foperand.length) {
                rangelookup = scf.OperandAsNumber(foperand);
                if (rangelookup.type.charAt(0) != "n") {
                    PushOperand("e#VALUE!", 0);
                    return;
                }
                if (foperand.length) {
                    scf.FunctionArgsError(fname, operand);
                    return 0;
                }
                rangelookup = rangelookup.value ? 1 : 0; // convert to 1 or 0
            }
        }
        lookupvalue.type = lookupvalue.type.charAt(0); // only deal with general type
        if (lookupvalue.type == "n") { // if number, make sure a number
            lookupvalue.value = lookupvalue.value - 0;
        }

        if (range.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }

        rangeinfo = scf.DecodeRangeParts(range.value, range.type);
        if (!rangeinfo) {
            PushOperand("e#REF!", 0);
            return;
        }

        c = 0;
        r = 0;
        cincr = 0;
        rincr = 0;
        if (fname == "HLOOKUP") {
            cincr = 1;
            if (offsetvalue > rangeinfo.nrows) {
                PushOperand("e#REF!", 0);
                return;
            }
        }
        else if (fname == "VLOOKUP") {
            rincr = 1;
            if (offsetvalue > rangeinfo.ncols) {
                PushOperand("e#REF!", 0);
                return;
            }
        }
        else if (fname == "MATCH") {
            if (rangeinfo.ncols > 1) {
                if (rangeinfo.nrows > 1) {
                    PushOperand("e#N/A", 0);
                    return;
                }
                cincr = 1;
            }
            else {
                rincr = 1;
            }
        }
        else {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        if (offsetvalue < 1 && fname != "MATCH") {
            PushOperand("e#VALUE!", 0);
            return 0;
        }

        previousOK; // if 1, previous test was <. If 2, also this one wasn't

        while (1) {
            cr = SocialCalc.crToCoord(rangeinfo.col1num + c, rangeinfo.row1num + r);
            cell = rangeinfo.sheetdata.GetAssuredCell(cr);
            value = cell.datavalue;
            valuetype = cell.valuetype ? cell.valuetype.charAt(0) : "b"; // only deal with general types
            if (valuetype == "n") {
                value = value - 0; // make sure number
            }
            if (rangelookup) { // rangelookup type 1 or -1: look for within brackets for matches
                if (lookupvalue.type == "n" && valuetype == "n") {
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                    if ((rangelookup > 0 && lookupvalue.value > value)
                        || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
                        previousOK = 1;
                        csave = c; // remember col and row of last OK
                        rsave = r;
                    }
                    else if (previousOK) { // last one was OK, this one isn't
                        previousOK = 2;
                        break;
                    }
                }

                else if (lookupvalue.type == "t" && valuetype == "t") {
                    value = typeof value == "string" ? value.toLowerCase() : "";
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                    if ((rangelookup > 0 && lookupvalue.value > value)
                        || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
                        previousOK = 1;
                        csave = c;
                        rsave = r;
                    }
                    else if (previousOK) { // last one was OK, this one isn't
                        previousOK = 2;
                        break;
                    }
                }
            }
            else { // exact value matches
                if (lookupvalue.type == "n" && valuetype == "n") {
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                }
                else if (lookupvalue.type == "t" && valuetype == "t") {
                    value = typeof value == "string" ? value.toLowerCase() : "";
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                }
            }

            r += rincr;
            c += cincr;
            if (r >= rangeinfo.nrows || c >= rangeinfo.ncols) { // end of range to check, no exact match
                if (previousOK) { // at least one could have been OK
                    previousOK = 2;
                    break;
                }
                PushOperand("e#N/A", 0);
                return;
            }
        }

        if (previousOK == 2) { // back to last OK
            r = rsave;
            c = csave;
        }

        if (fname == "MATCH") {
            value = c + r + 1; // only one may be <> 0
            valuetype = "n";
        }
        else {
            cr = SocialCalc.crToCoord(rangeinfo.col1num + c + (fname == "VLOOKUP" ? offsetvalue - 1 : 0), rangeinfo.row1num + r + (fname == "HLOOKUP" ? offsetvalue - 1 : 0));
            cell = rangeinfo.sheetdata.GetAssuredCell(cr);
            value = cell.datavalue;
            valuetype = cell.valuetype;
        }
        PushOperand(valuetype, value);

        return;

    }

    FunctionClassList['lookup']["HLOOKUP"] = [SocialCalcFormula.LookupFunctions, -3, "hlookup", "", "lookup"];
    FunctionClassList['lookup']["MATCH"] = [SocialCalcFormula.LookupFunctions, -2, "match", "", "lookup"];
    FunctionClassList['lookup']["VLOOKUP"] = [SocialCalcFormula.LookupFunctions, -3, "vlookup", "", "lookup"];

    /*
    #
    # INDEX(range, rownum, colnum)
    #
    */

    SocialCalcFormula.IndexFunction = function (fname, operand, foperand) {

        var range, sheetname, indexinfo, rowindex, colindex, result, resulttype;

        var scf = SocialCalcFormula;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        range = scf.TopOfStackValueAndType(foperand); // get range
        if (range.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        indexinfo = scf.DecodeRangeParts(range.value, range.type);
        if (indexinfo.sheetname) {
            sheetname = "!" + indexinfo.sheetname;
        }
        else {
            sheetname = "";
        }

        rowindex = {value: 0};
        colindex = {value: 0};

        if (foperand.length) { // look for row number
            rowindex = scf.OperandAsNumber(foperand);
            if (rowindex.type.charAt(0) != "n" || rowindex.value < 0) {
                PushOperand("e#VALUE!", 0);
                return;
            }
            if (foperand.length) { // look for col number
                colindex = scf.OperandAsNumber(foperand);
                if (colindex.type.charAt(0) != "n" || colindex.value < 0) {
                    PushOperand("e#VALUE!", 0);
                    return;
                }
                if (foperand.length) {
                    scf.FunctionArgsError(fname, operand);
                    return 0;
                }
            }
            else { // col number missing
                if (indexinfo.nrows == 1) { // if only one row, then rowindex is really colindex
                    colindex.value = rowindex.value;
                    rowindex.value = 0;
                }
            }
        }

        if (rowindex.value > indexinfo.nrows || colindex.value > indexinfo.ncols) {
            PushOperand("e#REF!", 0);
            return;
        }

        if (rowindex.value == 0) {
            if (colindex.value == 0) {
                if (indexinfo.nrows == 1 && indexinfo.ncols == 1) {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname;
                    resulttype = "coord";
                }
                else {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname + "|" +
                        SocialCalc.crToCoord(indexinfo.col1num + indexinfo.ncols - 1, indexinfo.row1num + indexinfo.nrows - 1) +
                        "|";
                    resulttype = "range";
                }
            }
            else {
                if (indexinfo.nrows == 1) {
                    result = SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num) + sheetname;
                    resulttype = "coord";
                }
                else {
                    result = SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num) + sheetname + "|" +
                        SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num + indexinfo.nrows - 1) +
                        "|";
                    resulttype = "range";
                }
            }
        }
        else {
            if (colindex.value == 0) {
                if (indexinfo.ncols == 1) {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num + rowindex.value - 1) + sheetname;
                    resulttype = "coord";
                }
                else {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num + rowindex.value - 1) + sheetname + "|" +
                        SocialCalc.crToCoord(indexinfo.col1num + indexinfo.ncols - 1, indexinfo.row1num + rowindex.value - 1) +
                        "|";
                    resulttype = "range";
                }
            }
            else {
                result = SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num + rowindex.value - 1) + sheetname;
                resulttype = "coord";
            }
        }

        PushOperand(resulttype, result);

        return;

    }

    FunctionClassList['lookup']["INDEX"] = [SocialCalcFormula.IndexFunction, -1, "index", "", "lookup"];

    /*
    #
    # COUNTIF(c1:c2,"criteria")
    # SUMIF(c1:c2,"criteria",[range2])
    #
    */

    SocialCalcFormula.CountifSumifFunctions = function (fname, operand, foperand) {

        var range, criteria, sumrange, f2operand, result, resulttype, value1, value2;
        var sum = 0;
        var resulttypesum = "";
        var count = 0;

        var scf = SocialCalcFormula;
        //var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        range = scf.TopOfStackValueAndType(foperand); // get range or coord
        criteria = scf.OperandAsText(foperand); // get criteria
        if (fname == "SUMIF") {
            if (foperand.length == 1) { // three arg form of SUMIF
                sumrange = scf.TopOfStackValueAndType(foperand);
            }
            else if (foperand.length == 0) { // two arg form
                sumrange = {value: range.value, type: range.type};
            }
            else {
                scf.FunctionArgsError(fname, operand);
                return 0;
            }
        }
        else {
            sumrange = {value: range.value, type: range.type};
        }

        if (criteria.type.charAt(0) == "n") {
            criteria.value = criteria.value + ""; // make text
        }
        else if (criteria.type.charAt(0) == "e") { // error
            criteria.value = null;
        }
        else if (criteria.type.charAt(0) == "b") { // blank here is undefined
            criteria.value = null;
        }

        if (range.type != "coord" && range.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }

        if (fname == "SUMIF" && sumrange.type != "coord" && sumrange.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }

        foperand.push(range);
        f2operand = []; // to allow for 3 arg form
        f2operand.push(sumrange);

        while (foperand.length) {
            value1 = scf.OperandValueAndType(foperand);
            value2 = scf.OperandValueAndType(f2operand);

            if (!scf.TestCriteria(value1.value, value1.type, criteria.value)) {
                continue;
            }

            count += 1;

            if (value2.type.charAt(0) == "n") {
                sum += value2.value - 0;
                resulttypesum = lookup_result_type(value2.type, resulttypesum || value2.type, typelookupplus);
            }
            else if (value2.type.charAt(0) == "e" && resulttypesum.charAt(0) != "e") {
                resulttypesum = value2.type;
            }
        }

        resulttypesum = resulttypesum || "n";

        if (fname == "SUMIF") {
            PushOperand(resulttypesum, sum);
        }
        else if (fname == "COUNTIF") {
            PushOperand("n", count);
        }

        return;

    }

    FunctionClassList['stat']["COUNTIF"] = [SocialCalcFormula.CountifSumifFunctions, 2, "rangec", "", "stat"];
    FunctionClassList['stat']["SUMIF"] = [SocialCalcFormula.CountifSumifFunctions, -2, "sumif", "", "stat"];

    /*
    #
    # IF(cond,truevalue,falsevalue)
    #
    */

    SocialCalcFormula.IfFunction = function (fname, operand, foperand) {

        var cond, t;

        cond = SocialCalcFormula.OperandValueAndType(foperand);
        t = cond.type.charAt(0);
        if (t != "n" && t != "b") {
            operand.push({type: "e#VALUE!", value: 0});
            return;
        }

        if (!cond.value) foperand.pop();
        operand.push(foperand.pop());
        if (cond.value) foperand.pop();

        return null;

    }

// Add to function list
    FunctionClassList['test']["IF"] = [SocialCalcFormula.IfFunction, 3, "iffunc", "", "test"];

    /*
    #
    # DATE(year,month,day)
    #
    */

    // SocialCalcFormula.DateFunction = function (fname, operand, foperand) {
    //
    //     var scf = SocialCalcFormula;
    //     var result = 0;
    //     var year = scf.OperandAsNumber(foperand);
    //     var month = scf.OperandAsNumber(foperand);
    //     var day = scf.OperandAsNumber(foperand);
    //     var resulttype = scf.LookupResultType(year.type, month.type, scf.TypeLookupTable.twoargnumeric);
    //     resulttype = scf.LookupResultType(resulttype, day.type, scf.TypeLookupTable.twoargnumeric);
    //     if (resulttype.charAt(0) == "n") {
    //         result = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(
    //             Math.floor(year.value), Math.floor(month.value), Math.floor(day.value)
    //         ) - SocialCalc.FormatNumber.datevalues.julian_offset;
    //         resulttype = "nd";
    //     }
    //     scf.PushOperand(operand, resulttype, result);
    //     return;
    //
    // }

    FunctionClassList['datetime']["DATE"] = [SocialCalcFormula.DateFunction, 3, "date", "", "datetime"];

    /*
    #
    # TIME(hour,minute,second)
    #
    */

    SocialCalcFormula.TimeFunction = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var hours = scf.OperandAsNumber(foperand);
        var minutes = scf.OperandAsNumber(foperand);
        var seconds = scf.OperandAsNumber(foperand);
        var resulttype = scf.LookupResultType(hours.type, minutes.type, scf.TypeLookupTable.twoargnumeric);
        resulttype = scf.LookupResultType(resulttype, seconds.type, scf.TypeLookupTable.twoargnumeric);
        if (resulttype.charAt(0) == "n") {
            result = ((hours.value * 60 * 60) + (minutes.value * 60) + seconds.value) / (24 * 60 * 60);
            resulttype = "nt";
        }
        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['datetime']["TIME"] = [SocialCalcFormula.TimeFunction, 3, "hms", "", "datetime"];

    /*
    #
    # DAY(date)
    # MONTH(date)
    # YEAR(date)
    # WEEKDAY(date, [type])
    #
    */

    // SocialCalcFormula.DMYFunctions = function (fname, operand, foperand) {
    //
    //     var ymd, dtype, doffset;
    //     var scf = SocialCalcFormula;
    //     var result = 0;
    //
    //     var datevalue = scf.OperandAsNumber(foperand);
    //     var resulttype = scf.LookupResultType(datevalue.type, datevalue.type, scf.TypeLookupTable.oneargnumeric);
    //
    //     if (resulttype.charAt(0) == "n") {
    //         ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(
    //             Math.floor(datevalue.value + SocialCalc.FormatNumber.datevalues.julian_offset));
    //         switch (fname) {
    //             case "DAY":
    //                 result = ymd.day;
    //                 break;
    //
    //             case "MONTH":
    //                 result = ymd.month;
    //                 break;
    //
    //             case "YEAR":
    //                 result = ymd.year;
    //                 break;
    //
    //             case "WEEKDAY":
    //                 dtype = {value: 1};
    //                 if (foperand.length) { // get type if present
    //                     dtype = scf.OperandAsNumber(foperand);
    //                     if (dtype.type.charAt(0) != "n" || dtype.value < 1 || dtype.value > 3) {
    //                         scf.PushOperand(operand, "e#VALUE!", 0);
    //                         return;
    //                     }
    //                     if (foperand.length) { // extra args
    //                         scf.FunctionArgsError(fname, operand);
    //                         return;
    //                     }
    //                 }
    //                 doffset = 6;
    //                 if (dtype.value > 1) {
    //                     doffset -= 1;
    //                 }
    //                 result = Math.floor(datevalue.value + doffset) % 7 + (dtype.value < 3 ? 1 : 0);
    //                 break;
    //         }
    //     }
    //
    //     scf.PushOperand(operand, resulttype, result);
    //     return;
    //
    // }

    FunctionClassList['datetime']["DAY"] = [SocialCalcFormula.DMYFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["MONTH"] = [SocialCalcFormula.DMYFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["YEAR"] = [SocialCalcFormula.DMYFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["WEEKDAY"] = [SocialCalcFormula.DMYFunctions, -1, "weekday", "", "datetime"];

    /*
    #
    # HOUR(datetime)
    # MINUTE(datetime)
    # SECOND(datetime)
    #
    */

    SocialCalcFormula.HMSFunctions = function (fname, operand, foperand) {

        var hours, minutes, seconds, fraction;
        var scf = SocialCalcFormula;
        var result = 0;

        var datetime = scf.OperandAsNumber(foperand);
        var resulttype = scf.LookupResultType(datetime.type, datetime.type, scf.TypeLookupTable.oneargnumeric);

        if (resulttype.charAt(0) == "n") {
            if (datetime.value < 0) {
                scf.PushOperand(operand, "e#NUM!", 0); // must be non-negative
                return;
            }
            fraction = datetime.value - Math.floor(datetime.value); // fraction of a day
            fraction *= 24;
            hours = Math.floor(fraction);
            fraction -= Math.floor(fraction);
            fraction *= 60;
            minutes = Math.floor(fraction);
            fraction -= Math.floor(fraction);
            fraction *= 60;
            seconds = Math.floor(fraction + (datetime.value >= 0 ? 0.5 : -0.5));
            if (fname == "HOUR") {
                result = hours;
            }
            else if (fname == "MINUTE") {
                result = minutes;
            }
            else if (fname == "SECOND") {
                result = seconds;
            }
        }

        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['datetime']["HOUR"] = [SocialCalcFormula.HMSFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["MINUTE"] = [SocialCalcFormula.HMSFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["SECOND"] = [SocialCalcFormula.HMSFunctions, 1, "v", "", "datetime"];

    /*
    #
    # EXACT(v1,v2)
    #
    */

    SocialCalcFormula.ExactFunction = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "nl";

        var value1 = scf.OperandValueAndType(foperand);
        var v1type = value1.type.charAt(0);
        var value2 = scf.OperandValueAndType(foperand);
        var v2type = value2.type.charAt(0);

        if (v1type == "t") {
            if (v2type == "t") {
                result = value1.value == value2.value ? 1 : 0;
            }
            else if (v2type == "b") {
                result = value1.value.length ? 0 : 1;
            }
            else if (v2type == "n") {
                result = value1.value == value2.value + "" ? 1 : 0;
            }
            else if (v2type == "e") {
                result = value2.value;
                resulttype = value2.type;
            }
            else {
                result = 0;
            }
        }
        else if (v1type == "n") {
            if (v2type == "n") {
                result = value1.value - 0 == value2.value - 0 ? 1 : 0;
            }
            else if (v2type == "b") {
                result = 0;
            }
            else if (v2type == "t") {
                result = value1.value + "" == value2.value ? 1 : 0;
            }
            else if (v2type == "e") {
                result = value2.value;
                resulttype = value2.type;
            }
            else {
                result = 0;
            }
        }
        else if (v1type == "b") {
            if (v2type == "t") {
                result = value2.value.length ? 0 : 1;
            }
            else if (v2type == "b") {
                result = 1;
            }
            else if (v2type == "n") {
                result = 0;
            }
            else if (v2type == "e") {
                result = value2.value;
                resulttype = value2.type;
            }
            else {
                result = 0;
            }
        }
        else if (v1type == "e") {
            result = value1.value;
            resulttype = value1.type;
        }

        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['text']["EXACT"] = [SocialCalcFormula.ExactFunction, 2, "", "", "text"];

    /*
    #
    # FIND(key,string,[start])
    # LEFT(string,[length])
    # LEN(string)
    # LOWER(string)
    # MID(string,start,length)
    # PROPER(string)
    # REPLACE(string,start,length,new)
    # REPT(string,count)
    # RIGHT(string,[length])
    # SUBSTITUTE(string,old,new,[which])
    # TRIM(string)
    # UPPER(string)
    #
    */

// SocialCalcFormula.ArgList has an array for each function, one entry for each possible arg (up to max).
// Min args are specified in SocialCalcFormula.FunctionList.
// If array element is 1 then it's a text argument, if it's 0 then it's numeric, if -1 then just get whatever's there
// Text values are manipulated as UTF-8, converting from and back to byte strings

    SocialCalcFormula.ArgList = {
        FIND: [1, 1, 0],
        LEFT: [1, 0],
        LEN: [1],
        LOWER: [1],
        MID: [1, 0, 0],
        PROPER: [1],
        REPLACE: [1, 0, 0, 1],
        REPT: [1, 0],
        RIGHT: [1, 0],
        SUBSTITUTE: [1, 1, 1, 0],
        TRIM: [1],
        UPPER: [1]
    };

    SocialCalcFormula.StringFunctions = function (fname, operand, foperand) {
        //console.log(fname)
        var i, value, offset, len, start, count;
        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "e#VALUE!";

        var numargs = foperand.length;
        var argdef = scf.ArgList[fname];
        var operand_value = [];
        var operand_type = [];

        for (i = 1; i <= numargs; i++) { // go through each arg, get value and type, and check for errors
            if (i > argdef.length) { // too many args
                scf.FunctionArgsError(fname, operand);
                return;
            }
            if (argdef[i - 1] == 0) {
                value = scf.OperandAsNumber(foperand);
            }
            else if (argdef[i - 1] == 1) {
                value = scf.OperandAsText(foperand);
            }
            else if (argdef[i - 1] == -1) {
                value = scf.OperandValueAndType(foperand);
            }
            operand_value[i] = value.value;
            operand_type[i] = value.type;
            if (value.type.charAt(0) == "e") {
                scf.PushOperand(operand, value.type, result);
                return;
            }
        }

        switch (fname) {
            case "FIND":
                offset = operand_type[3] ? operand_value[3] - 1 : 0;
                if (offset < 0) {
                    result = "Start is before string"; // !! not displayed, no need to translate
                }
                else {
                    result = operand_value[2].indexOf(operand_value[1], offset); // (null string matches first char)
                    if (result >= 0) {
                        result += 1;
                        resulttype = "n";
                    }
                    else {
                        result = "Not found"; // !! not displayed, error is e#VALUE!
                    }
                }
                break;

            case "LEFT":
                len = operand_type[2] ? operand_value[2] - 0 : 1;
                if (len < 0) {
                    result = "Negative length";
                }
                else {
                    result = operand_value[1].substring(0, len);
                    resulttype = "t";
                }
                break;

            case "LEN":
                result = (operand_value[1] + '').length;
                resulttype = "n";
                break;

            case "LOWER":
                result = operand_value[1].toLowerCase();
                resulttype = "t";
                break;

            case "MID":
                start = operand_value[2] - 0;
                len = operand_value[3] - 0;
                if (len < 1 || start < 1) {
                    result = "Bad arguments";
                }
                else {
                    result = operand_value[1].substring(start - 1, start + len - 1);
                    resulttype = "t";
                }
                break;

            case "PROPER":
                result = operand_value[1].replace(/\b\w+\b/g, function (word) {
                    return word.substring(0, 1).toUpperCase() +
                        word.substring(1);
                }); // uppercase first character of words (see JavaScript, Flanagan, 5th edition, page 704)
                resulttype = "t";
                break;

            case "REPLACE":
                start = operand_value[2] - 0;
                len = operand_value[3] - 0;
                if (len < 0 || start < 1) {
                    result = "Bad arguments";
                }
                else {
                    result = operand_value[1].substring(0, start - 1) + operand_value[4] +
                        operand_value[1].substring(start - 1 + len);
                    resulttype = "t";
                }
                break;

            case "REPT":
                count = operand_value[2] - 0;
                if (count < 0) {
                    result = "Negative count";
                }
                else {
                    result = "";
                    for (; count > 0; count--) {
                        result += operand_value[1];
                    }
                    resulttype = "t";
                }
                break;

            case "RIGHT":
                len = operand_type[2] ? operand_value[2] - 0 : 1;
                if (len < 0) {
                    result = "Negative length";
                }
                else {
                    result = operand_value[1].slice(-len);
                    resulttype = "t";
                }
                break;

            case "SUBSTITUTE":
                fulltext = operand_value[1];
                oldtext = operand_value[2];
                newtext = operand_value[3];
                if (operand_value[4] != null) {
                    which = operand_value[4] - 0;
                    if (which <= 0) {
                        result = "Non-positive instance number";
                        break;
                    }
                }
                else {
                    which = 0;
                }
                count = 0;
                oldpos = 0;
                result = "";
                while (true) {
                    pos = fulltext.indexOf(oldtext, oldpos);
                    if (pos >= 0) {
                        count++; //!!!!!! old test just in case: if (count>1000) {alert(pos); break;}
                        result += fulltext.substring(oldpos, pos);
                        if (which == 0) {
                            result += newtext; // substitute
                        }
                        else if (which == count) {
                            result += newtext + fulltext.substring(pos + oldtext.length);
                            break;
                        }
                        else {
                            result += oldtext; // leave as was
                        }
                        oldpos = pos + oldtext.length;
                    }
                    else { // no more
                        result += fulltext.substring(oldpos);
                        break;
                    }
                }
                resulttype = "t";
                break;

            case "TRIM":
                result = operand_value[1];
                result = result.replace(/^ */, "");
                result = result.replace(/ *$/, "");
                result = result.replace(/ +/g, " ");
                resulttype = "t";
                break;

            case "UPPER":
                result = operand_value[1].toUpperCase();
                resulttype = "t";
                break;

        }

        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['text']["FIND"] = [SocialCalcFormula.StringFunctions, -2, "find", "", "text"];
    FunctionClassList['text']["LEFT"] = [SocialCalcFormula.StringFunctions, -2, "tc", "", "text"];
    FunctionClassList['text']["LEN"] = [SocialCalcFormula.StringFunctions, 1, "txt", "", "text"];
    FunctionClassList['text']["LOWER"] = [SocialCalcFormula.StringFunctions, 1, "txt", "", "text"];
    FunctionClassList['text']["MID"] = [SocialCalcFormula.StringFunctions, 3, "mid", "", "text"];
    FunctionClassList['text']["PROPER"] = [SocialCalcFormula.StringFunctions, 1, "v", "", "text"];
    FunctionClassList['text']["REPLACE"] = [SocialCalcFormula.StringFunctions, 4, "replace", "", "text"];
    FunctionClassList['text']["REPT"] = [SocialCalcFormula.StringFunctions, 2, "tc", "", "text"];
    FunctionClassList['text']["RIGHT"] = [SocialCalcFormula.StringFunctions, -1, "tc", "", "text"];
    FunctionClassList['text']["SUBSTITUTE"] = [SocialCalcFormula.StringFunctions, -3, "subs", "", "text"];
    FunctionClassList['text']["TRIM"] = [SocialCalcFormula.StringFunctions, 1, "v", "", "text"];
    FunctionClassList['text']["UPPER"] = [SocialCalcFormula.StringFunctions, 1, "v", "", "text"];

    /*
    #
    # is_functions:
    #
    # ISBLANK(value)
    # ISERR(value)
    # ISERROR(value)
    # ISLOGICAL(value)
    # ISNA(value)
    # ISNONTEXT(value)
    # ISNUMBER(value)
    # ISTEXT(value)
    #
    */

    SocialCalcFormula.IsFunctions = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "nl";

        var value = scf.OperandValueAndType(foperand);
        var t = value.type.charAt(0);

        switch (fname) {

            case "ISBLANK":
                result = value.type == "b" ? 1 : 0;
                break;

            case "ISERR":
                result = t == "e" ? (value.type == "e#N/A" ? 0 : 1) : 0;
                break;

            case "ISERROR":
                result = t == "e" ? 1 : 0;
                break;

            case "ISLOGICAL":
                result = value.type == "nl" ? 1 : 0;
                break;

            case "ISNA":
                result = value.type == "e#N/A" ? 1 : 0;
                break;

            case "ISNONTEXT":
                result = t == "t" ? 0 : 1;
                break;

            case "ISNUMBER":
                result = t == "n" ? 1 : 0;
                break;

            case "ISTEXT":
                result = t == "t" ? 1 : 0;
                break;
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['test']["ISBLANK"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISERR"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISERROR"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISLOGICAL"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISNA"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISNONTEXT"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISNUMBER"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISTEXT"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];

    /*
    #
    # ntv_functions:
    #
    # N(value)
    # T(value)
    # VALUE(value)
    #
    */

    SocialCalcFormula.NTVFunctions = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "e#VALUE!";

        var value = scf.OperandValueAndType(foperand);
        var t = value.type.charAt(0);

        switch (fname) {

            case "N":
                result = t == "n" ? value.value - 0 : 0;
                resulttype = "n";
                break;

            case "T":
                result = t == "t" ? value.value + "" : "";
                resulttype = "t";
                break;

            case "VALUE":
                if (t == "n" || t == "b") {
                    result = value.value || 0;
                    resulttype = "n";
                }
                else if (t == "t") {
                    value = scf.DetermineValueType(value.value);
                    if (value.type.charAt(0) != "n") {
                        result = 0;
                        resulttype = "e#VALUE!";
                    }
                    else {
                        result = value.value - 0;
                        resulttype = "n";
                    }
                }
                break;
        }

        if (t == "e") { // error trumps
            resulttype = value.type;
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['math']["N"] = [SocialCalcFormula.NTVFunctions, 1, "v", "", "math"];
    FunctionClassList['text']["T"] = [SocialCalcFormula.NTVFunctions, 1, "v", "", "text"];
    FunctionClassList['text']["VALUE"] = [SocialCalcFormula.NTVFunctions, 1, "v", "", "text"];

    /*
    #
    # ABS(value)
    # ACOS(value)
    # ASIN(value)
    # ATAN(value)
    # COS(value)
    # DEGREES(value)
    # EVEN(value)
    # EXP(value)
    # FACT(value)
    # INT(value)
    # LN(value)
    # LOG10(value)
    # ODD(value)
    # RADIANS(value)
    # SIN(value)
    # SQRT(value)
    # TAN(value)
    #
    */

    SocialCalcFormula.Math1Functions = function (fname, operand, foperand) {

        var v1, value, f;
        var result = {};

        var scf = SocialCalcFormula;

        v1 = scf.OperandAsNumber(foperand);
        value = v1.value;
        result.type = scf.LookupResultType(v1.type, v1.type, scf.TypeLookupTable.oneargnumeric);

        if (result.type == "n") {
            switch (fname) {
                case "ABS":
                    value = Math.abs(value);
                    break;

                case "ACOS":
                    if (value >= -1 && value <= 1) {
                        value = Math.acos(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;

                case "ASIN":
                    if (value >= -1 && value <= 1) {
                        value = Math.asin(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;

                case "ATAN":
                    value = Math.atan(value);
                    break;

                case "COS":
                    value = Math.cos(value);
                    break;

                case "DEGREES":
                    value = value * 180 / Math.PI;
                    break;

                case "EVEN":
                    value = value < 0 ? -value : value;
                    if (value != Math.floor(value)) {
                        value = Math.floor(value + 1) + (Math.floor(value + 1) % 2);
                    }
                    else { // integer
                        value = value + (value % 2);
                    }
                    if (v1.value < 0) value = -value;
                    break;

                case "EXP":
                    value = Math.exp(value);
                    break;

                case "FACT":
                    f = 1;
                    value = Math.floor(value);
                    for (; value > 0; value--) {
                        f *= value;
                    }
                    value = f;
                    break;

                case "INT":
                    value = Math.floor(value); // spreadsheet INT is floor(), not int()
                    break;

                case "LN":
                    if (value <= 0) {
                        result.type = "e#NUM!";
                        result.error = SocialCalc.Constants.s_sheetfunclnarg;
                    }
                    value = Math.log(value);
                    break;

                case "LOG10":
                    if (value <= 0) {
                        result.type = "e#NUM!";
                        result.error = SocialCalc.Constants.s_sheetfunclog10arg;
                    }
                    value = Math.log(value) / Math.log(10);
                    break;

                case "ODD":
                    value = value < 0 ? -value : value;
                    if (value != Math.floor(value)) {
                        value = Math.floor(value + 1) + (1 - (Math.floor(value + 1) % 2));
                    }
                    else { // integer
                        value = value + (1 - (value % 2));
                    }
                    if (v1.value < 0) value = -value;
                    break;

                case "RADIANS":
                    value = value * Math.PI / 180;
                    break;

                case "SIN":
                    value = Math.sin(value);
                    break;

                case "SQRT":
                    if (value >= 0) {
                        value = Math.sqrt(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;

                case "TAN":
                    if (Math.cos(value) != 0) {
                        value = Math.tan(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;
            }
        }

        result.value = value;
        operand.push(result);

        return null;

    }

// Add to function list
    FunctionClassList['math']["ABS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ACOS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ASIN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ATAN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["COS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["DEGREES"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["EVEN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["EXP"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["FACT"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["INT"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["LN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["LOG10"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ODD"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["RADIANS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["SIN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["SQRT"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["TAN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];


    /*
    #
    # ATAN2(x, y)
    # MOD(a, b)
    # POWER(a, b)
    # TRUNC(value, precision)
    #
    */

    SocialCalcFormula.Math2Functions = function (fname, operand, foperand) {

        var xval, yval, value, quotient, decimalscale, i;
        var result = {};

        var scf = SocialCalcFormula;

        xval = scf.OperandAsNumber(foperand);
        yval = scf.OperandAsNumber(foperand);
        value = 0;
        result.type = scf.LookupResultType(xval.type, yval.type, scf.TypeLookupTable.twoargnumeric);

        if (result.type == "n") {
            switch (fname) {
                case "ATAN2":
                    if (xval.value == 0 && yval.value == 0) {
                        result.type = "e#DIV/0!";
                    }
                    else {
                        result.value = Math.atan2(yval.value, xval.value);
                    }
                    break;

                case "POWER":
                    result.value = Math.pow(xval.value, yval.value);
                    if (isNaN(result.value)) {
                        result.value = 0;
                        result.type = "e#NUM!";
                    }
                    break;

                case "MOD": // en.wikipedia.org/wiki/Modulo_operation, etc.
                    if (yval.value == 0) {
                        result.type = "e#DIV/0!";
                    }
                    else {
                        quotient = xval.value / yval.value;
                        quotient = Math.floor(quotient);
                        result.value = xval.value - (quotient * yval.value);
                    }
                    break;

                case "TRUNC":
                    decimalscale = 1; // cut down to required number of decimal digits
                    if (yval.value >= 0) {
                        yval.value = Math.floor(yval.value);
                        for (i = 0; i < yval.value; i++) {
                            decimalscale *= 10;
                        }
                        result.value = Math.floor(Math.abs(xval.value) * decimalscale) / decimalscale;
                    }
                    else if (yval.value < 0) {
                        yval.value = Math.floor(-yval.value);
                        for (i = 0; i < yval.value; i++) {
                            decimalscale *= 10;
                        }
                        result.value = Math.floor(Math.abs(xval.value) / decimalscale) * decimalscale;
                    }
                    if (xval.value < 0) {
                        result.value = -result.value;
                    }
            }
        }

        operand.push(result);

        return null;

    }

// Add to function list
    FunctionClassList['math']["ATAN2"] = [SocialCalcFormula.Math2Functions, 2, "xy", "", "math"];
    FunctionClassList['math']["MOD"] = [SocialCalcFormula.Math2Functions, 2, "", "", "math"];
    FunctionClassList['math']["POWER"] = [SocialCalcFormula.Math2Functions, 2, "", "", "math"];
    FunctionClassList['math']["TRUNC"] = [SocialCalcFormula.Math2Functions, 2, "valpre", "", "math"];

    /*
    #
    # LOG(value,[base])
    #
    */

    SocialCalcFormula.LogFunction = function (fname, operand, foperand) {

        var value, value2;
        var result = {};

        var scf = SocialCalcFormula;

        result.value = 0;

        value = scf.OperandAsNumber(foperand);
        result.type = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);
        if (foperand.length == 1) {
            value2 = scf.OperandAsNumber(foperand);
            if (value2.type.charAt(0) != "n" || value2.value <= 0) {
                scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogsecondarg);
                return 0;
            }
        }
        else if (foperand.length != 0) {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        else {
            value2 = {value: Math.E, type: "n"};
        }

        if (result.type == "n") {
            if (value.value <= 0) {
                scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogfirstarg);
                return 0;
            }
            result.value = Math.log(value.value) / Math.log(value2.value);
        }

        operand.push(result);

        return;

    }

    FunctionClassList['math']["LOG"] = [SocialCalcFormula.LogFunction, -1, "log", "", "math"];


    /*
    #
    # ROUND(value,[precision])
    #
    */

    SocialCalcFormula.RoundFunction = function (fname, operand, foperand) {

        var value2, decimalscale, scaledvalue, i;

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "e#VALUE!";

        var value = scf.OperandValueAndType(foperand);
        var resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);

        if (foperand.length == 1) {
            value2 = scf.OperandValueAndType(foperand);
            if (value2.type.charAt(0) != "n") {
                scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncroundsecondarg);
                return 0;
            }
        }
        else if (foperand.length != 0) {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        else {
            value2 = {value: 0, type: "n"}; // if no second arg, assume 0 for simple round
        }

        if (resulttype == "n") {
            value2.value = value2.value - 0;
            if (value2.value == 0) {
                result = Math.round(value.value);
            }
            else if (value2.value > 0) {
                decimalscale = 1; // cut down to required number of decimal digits
                value2.value = Math.floor(value2.value);
                for (i = 0; i < value2.value; i++) {
                    decimalscale *= 10;
                }
                scaledvalue = Math.round(value.value * decimalscale);
                result = scaledvalue / decimalscale;
            }
            else if (value2.value < 0) {
                decimalscale = 1; // cut down to required number of decimal digits
                value2.value = Math.floor(-value2.value);
                for (i = 0; i < value2.value; i++) {
                    decimalscale *= 10;
                }
                scaledvalue = Math.round(value.value / decimalscale);
                result = scaledvalue * decimalscale;
            }
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['math']["ROUND"] = [SocialCalcFormula.RoundFunction, -1, "vp", "", "math"];

    /*
    #
    # AND(v1,c1:c2,...)
    # OR(v1,c1:c2,...)
    #
    */

    SocialCalcFormula.AndOrFunctions = function (fname, operand, foperand) {

        var value1, result;

        var scf = SocialCalcFormula;
        var resulttype = "";

        if (fname == "AND") {
            result = 1;
        }
        else if (fname == "OR") {
            result = 0;
        }

        while (foperand.length) {
            value1 = scf.OperandValueAndType(foperand);
            if (value1.type.charAt(0) == "n") {
                value1.value = value1.value - 0;
                if (fname == "AND") {
                    result = value1.value != 0 ? result : 0;
                }
                else if (fname == "OR") {
                    result = value1.value != 0 ? 1 : result;
                }
                resulttype = scf.LookupResultType(value1.type, resulttype || "nl", scf.TypeLookupTable.propagateerror);
            }
            else if (value1.type.charAt(0) == "e" && resulttype.charAt(0) != "e") {
                resulttype = value1.type;
            }
        }
        if (resulttype.length < 1) {
            resulttype = "e#VALUE!";
            result = 0;
        }
        //console.log(scf)
        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['test']["AND"] = [SocialCalcFormula.AndOrFunctions, -1, "vn", "", "test"];
    FunctionClassList['test']["OR"] = [SocialCalcFormula.AndOrFunctions, -1, "vn", "", "test"];

    /*
    #
    # NOT(value)
    #
    */

    SocialCalcFormula.NotFunction = function (fname, operand, foperand) {

        var result = 0;
        var scf = SocialCalcFormula;
        var value = scf.OperandValueAndType(foperand);
        var resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.propagateerror);

        if (value.type.charAt(0) == "n" || value.type == "b") {
            result = value.value - 0 != 0 ? 0 : 1; // do the "not" operation
            resulttype = "nl";
        }
        else if (value.type.charAt(0) == "t") {
            resulttype = "e#VALUE!";
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['test']["NOT"] = [SocialCalcFormula.NotFunction, 1, "v", "", "test"];

    /*
    #
    # CHOOSE(index,value1,value2,...)
    #
    */

    SocialCalcFormula.ChooseFunction = function (fname, operand, foperand) {

        var resulttype, count, value1;
        var result = 0;
        var scf = SocialCalcFormula;

        var cindex = scf.OperandAsNumber(foperand);

        if (cindex.type.charAt(0) != "n") {
            cindex.value = 0;
        }
        cindex.value = Math.floor(cindex.value);

        count = 0;
        while (foperand.length) {
            value1 = scf.TopOfStackValueAndType(foperand);
            count += 1;
            if (cindex.value == count) {
                result = value1.value;
                resulttype = value1.type;
                break;
            }
        }
        if (resulttype) { // found something
            scf.PushOperand(operand, resulttype, result);
        }
        else {
            scf.PushOperand(operand, "e#VALUE!", 0);
        }

        return;

    }

    FunctionClassList['lookup']["CHOOSE"] = [SocialCalcFormula.ChooseFunction, -2, "choose", "", "lookup"];

    /*
    #
    # COLUMNS(c1:c2)
    # ROWS(c1:c2)
    #
    */

    SocialCalcFormula.ColumnsRowsFunctions = function (fname, operand, foperand) {

        var resulttype, rangeinfo;
        var result = 0;
        var scf = SocialCalcFormula;

        var value1 = scf.TopOfStackValueAndType(foperand);

        if (value1.type == "coord") {
            result = 1;
            resulttype = "n";
        }

        else if (value1.type == "range") {
            rangeinfo = scf.DecodeRangeParts(value1.value);
            if (fname == "COLUMNS") {
                result = rangeinfo.ncols;
            }
            else if (fname == "ROWS") {
                result = rangeinfo.nrows;
            }
            resulttype = "n";
        }
        else {
            result = 0;
            resulttype = "e#VALUE!";
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['lookup']["COLUMNS"] = [SocialCalcFormula.ColumnsRowsFunctions, 1, "range", "", "lookup"];
    FunctionClassList['lookup']["ROWS"] = [SocialCalcFormula.ColumnsRowsFunctions, 1, "range", "", "lookup"];


    /*
    #
    # FALSE()
    # NA()
    # NOW()
    # PI()
    # TODAY()
    # TRUE()
    #
    */

    SocialCalcFormula.ZeroArgFunctions = function (fname, operand, foperand) {

        var startval, tzoffset, start_1_1_1970, seconds_in_a_day, nowdays;
        var result = {value: 0};

        switch (fname) {
            case "FALSE":
                result.type = "nl";
                result.value = 0;
                break;

            case "NA":
                result.type = "e#N/A";
                break;

            case "NOW":
                startval = new Date();
                tzoffset = startval.getTimezoneOffset();
                startval = startval.getTime() / 1000; // convert to seconds
                start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
                seconds_in_a_day = 24 * 60 * 60;
                nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset / (24 * 60);
                result.value = nowdays;
                result.type = "ndt";
                //SocialCalcFormula.FreshnessInfo.volatile.NOW = true; // remember
                break;

            case "PI":
                result.type = "n";
                result.value = Math.PI;
                break;

            case "TODAY":
                startval = new Date();
                tzoffset = startval.getTimezoneOffset();
                startval = startval.getTime() / 1000; // convert to seconds
                start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
                seconds_in_a_day = 24 * 60 * 60;
                nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset / (24 * 60);
                result.value = Math.floor(nowdays);
                result.type = "nd";
                //SocialCalcFormula.FreshnessInfo.volatile.TODAY = true; // remember
                break;

            case "TRUE":
                result.type = "nl";
                result.value = 1;
                break;

        }

        operand.push(result);

        return null;

    }

// Add to function list
    FunctionClassList['test']["FALSE"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "test"];
    FunctionClassList['test']["NA"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "test"];
    FunctionClassList['datetime']["NOW"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "datetime"];
    FunctionClassList['math']["PI"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "math"];
    FunctionClassList['datetime']["TODAY"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "datetime"];
    FunctionClassList['test']["TRUE"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "test"];

//
// * * * * * FINANCIAL FUNCTIONS * * * * *
//

    /*
    #
    # DDB(cost,salvage,lifetime,period,[method])
    #
    # Depreciation, method defaults to 2 for double-declining balance
    # See: http://en.wikipedia.org/wiki/Depreciation
    #
    */

    SocialCalcFormula.DDBFunction = function (fname, operand, foperand) {

        var method, depreciation, accumulateddepreciation, i;
        var scf = SocialCalcFormula;

        var cost = scf.OperandAsNumber(foperand);
        var salvage = scf.OperandAsNumber(foperand);
        var lifetime = scf.OperandAsNumber(foperand);
        var period = scf.OperandAsNumber(foperand);

        if (scf.CheckForErrorValue(operand, cost)) return;
        if (scf.CheckForErrorValue(operand, salvage)) return;
        if (scf.CheckForErrorValue(operand, lifetime)) return;
        if (scf.CheckForErrorValue(operand, period)) return;

        if (lifetime.value < 1) {
            scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncddblife);
            return 0;
        }

        method = {value: 2, type: "n"};
        if (foperand.length > 0) {
            method = scf.OperandAsNumber(foperand);
        }
        if (foperand.length != 0) {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        if (scf.CheckForErrorValue(operand, method)) return;

        depreciation = 0; // calculated for each period
        accumulateddepreciation = 0; // accumulated by adding each period's

        for (i = 1; i <= period.value - 0 && i <= lifetime.value; i++) { // calculate for each period based on net from previous
            depreciation = (cost.value - accumulateddepreciation) * (method.value / lifetime.value);
            if (cost.value - accumulateddepreciation - depreciation < salvage.value) { // don't go lower than salvage value
                depreciation = cost.value - accumulateddepreciation - salvage.value;
            }
            accumulateddepreciation += depreciation;
        }

        scf.PushOperand(operand, 'n$', depreciation);

        return;

    }

    FunctionClassList['financial']["DDB"] = [SocialCalcFormula.DDBFunction, -4, "ddb", "", "financial"];

    /*
    #
    # SLN(cost,salvage,lifetime)
    #
    # Depreciation for each period by straight-line method
    # See: http://en.wikipedia.org/wiki/Depreciation
    #
    */

    SocialCalcFormula.SLNFunction = function (fname, operand, foperand) {

        var depreciation;
        var scf = SocialCalcFormula;

        var cost = scf.OperandAsNumber(foperand);
        var salvage = scf.OperandAsNumber(foperand);
        var lifetime = scf.OperandAsNumber(foperand);

        if (scf.CheckForErrorValue(operand, cost)) return;
        if (scf.CheckForErrorValue(operand, salvage)) return;
        if (scf.CheckForErrorValue(operand, lifetime)) return;

        if (lifetime.value < 1) {
            scf.FunctionSpecificError(fname, operand, "e#NUM!", scf.scc.s_sheetfuncslnlife);
            return 0;
        }

        depreciation = (cost.value - salvage.value) / lifetime.value;

        scf.PushOperand(operand, 'n$', depreciation);

        return;

    }

    FunctionClassList['financial']["SLN"] = [SocialCalcFormula.SLNFunction, 3, "csl", "", "financial"];

    /*
    #
    # SYD(cost,salvage,lifetime,period)
    #
    # Depreciation by Sum of Year's Digits method
    #
    */

    SocialCalcFormula.SYDFunction = function (fname, operand, foperand) {

        var depreciation, sumperiods;
        var scf = SocialCalcFormula;

        var cost = scf.OperandAsNumber(foperand);
        var salvage = scf.OperandAsNumber(foperand);
        var lifetime = scf.OperandAsNumber(foperand);
        var period = scf.OperandAsNumber(foperand);

        if (scf.CheckForErrorValue(operand, cost)) return;
        if (scf.CheckForErrorValue(operand, salvage)) return;
        if (scf.CheckForErrorValue(operand, lifetime)) return;
        if (scf.CheckForErrorValue(operand, period)) return;

        if (lifetime.value < 1 || period.value <= 0) {
            scf.PushOperand(operand, "e#NUM!", 0);
            return 0;
        }

        sumperiods = ((lifetime.value + 1) * lifetime.value) / 2; // add up 1 through lifetime
        depreciation = (cost.value - salvage.value) * (lifetime.value - period.value + 1) / sumperiods; // calc depreciation

        scf.PushOperand(operand, 'n$', depreciation);

        return;

    }

    FunctionClassList['financial']["SYD"] = [SocialCalcFormula.SYDFunction, 4, "cslp", "", "financial"];

    /*
    #
    # FV(rate, n, payment, [pv, [paytype]])
    # NPER(rate, payment, pv, [fv, [paytype]])
    # PMT(rate, n, pv, [fv, [paytype]])
    # PV(rate, n, payment, [fv, [paytype]])
    # RATE(n, payment, pv, [fv, [paytype, [guess]]])
    #
    # Following the Open Document Format formula specification:
    #
    #    PV = - Fv - (Payment * Nper) [if rate equals 0]
    #    Pv*(1+Rate)^Nper + Payment * (1 + Rate*PaymentType) * ( (1+Rate)^nper -1)/Rate + Fv = 0
    #
    # For each function, the formulas are solved for the appropriate value (transformed using
    # basic algebra).
    #
    */

    SocialCalcFormula.InterestFunctions = function (fname, operand, foperand) {

        var resulttype, result, dval, eval, fval;
        var pv, fv, rate, n, payment, paytype, guess, part1, part2, part3, part4, part5;
        var olddelta, maxloop, tries, deltaepsilon, rate, oldrate, m;

        var scf = SocialCalcFormula;

        var aval = scf.OperandAsNumber(foperand);
        var bval = scf.OperandAsNumber(foperand);
        var cval = scf.OperandAsNumber(foperand);

        resulttype = scf.LookupResultType(aval.type, bval.type, scf.TypeLookupTable.twoargnumeric);
        resulttype = scf.LookupResultType(resulttype, cval.type, scf.TypeLookupTable.twoargnumeric);
        if (foperand.length) { // optional arguments
            dval = scf.OperandAsNumber(foperand);
            resulttype = scf.LookupResultType(resulttype, dval.type, scf.TypeLookupTable.twoargnumeric);
            if (foperand.length) { // optional arguments
                eval = scf.OperandAsNumber(foperand);
                resulttype = scf.LookupResultType(resulttype, eval.type, scf.TypeLookupTable.twoargnumeric);
                if (foperand.length) { // optional arguments
                    if (fname != "RATE") { // only rate has 6 possible args
                        scf.FunctionArgsError(fname, operand);
                        return 0;
                    }
                    fval = scf.OperandAsNumber(shfoperand);
                    resulttype = scf.LookupResultType(resulttype, fval.type, scf.TypeLookupTable.twoargnumeric);
                }
            }
        }

        if (resulttype == "n") {
            switch (fname) {
                case "FV": // FV(rate, n, payment, [pv, [paytype]])
                    rate = aval.value;
                    n = bval.value;
                    payment = cval.value;
                    pv = dval != null ? dval.value : 0; // get value if present, or use default
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (rate == 0) { // simple calculation if no interest
                        fv = -pv - (payment * n);
                    }
                    else {
                        fv = -(pv * Math.pow(1 + rate, n) + payment * (1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate);
                    }
                    result = fv;
                    resulttype = 'n$';
                    break;

                case "NPER": // NPER(rate, payment, pv, [fv, [paytype]])
                    rate = aval.value;
                    payment = bval.value;
                    pv = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (rate == 0) { // simple calculation if no interest
                        if (payment == 0) {
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                        n = (pv + fv) / (-payment);
                    }
                    else {
                        part1 = payment * (1 + rate * paytype) / rate;
                        part2 = pv + part1;
                        if (part2 == 0 || rate <= -1) {
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                        part3 = (part1 - fv) / part2;
                        if (part3 <= 0) {
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                        part4 = Math.log(part3);
                        part5 = Math.log(1 + rate); // rate > -1
                        n = part4 / part5;
                    }
                    result = n;
                    resulttype = 'n';
                    break;

                case "PMT": // PMT(rate, n, pv, [fv, [paytype]])
                    rate = aval.value;
                    n = bval.value;
                    pv = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (n == 0) {
                        scf.PushOperand(operand, "e#NUM!", 0);
                        return;
                    }
                    else if (rate == 0) { // simple calculation if no interest
                        payment = (fv - pv) / n;
                    }
                    else {
                        payment = (0 - fv - pv * Math.pow(1 + rate, n)) / ((1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate);
                    }
                    result = payment;
                    resulttype = 'n$';
                    break;

                case "PV": // PV(rate, n, payment, [fv, [paytype]])
                    rate = aval.value;
                    n = bval.value;
                    payment = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (rate == -1) {
                        scf.PushOperand(operand, "e#DIV/0!", 0);
                        return;
                    }
                    else if (rate == 0) { // simple calculation if no interest
                        pv = -fv - (payment * n);
                    }
                    else {
                        pv = (-fv - payment * (1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate) / (Math.pow(1 + rate, n));
                    }
                    result = pv;
                    resulttype = 'n$';
                    break;

                case "RATE": // RATE(n, payment, pv, [fv, [paytype, [guess]]])
                    n = aval.value;
                    payment = bval.value;
                    pv = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    guess = fval != null ? fval.value : 0.1;

                    // rate is calculated by repeated approximations
                    // The deltas are used to calculate new guesses

                    maxloop = 100;
                    tries = 0;
                    delta = 1;
                    epsilon = 0.0000001; // this is close enough
                    rate = guess || 0.00000001; // zero is not allowed
                    while ((delta >= 0 ? delta : -delta) > epsilon && (rate != oldrate)) {
                        delta = fv + pv * Math.pow(1 + rate, n) + payment * (1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate;
                        if (olddelta != null) {
                            m = (delta - olddelta) / (rate - oldrate) || .001; // get slope (not zero)
                            oldrate = rate;
                            rate = rate - delta / m; // look for zero crossing
                            olddelta = delta;
                        }
                        else { // first time - no old values
                            oldrate = rate;
                            rate = 1.1 * rate;
                            olddelta = delta;
                        }
                        tries++;
                        if (tries >= maxloop) { // didn't converge yet
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                    }
                    result = rate;
                    resulttype = 'n%';
                    break;
            }
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['financial']["FV"] = [SocialCalcFormula.InterestFunctions, -3, "fv", "", "financial"];
    FunctionClassList['financial']["NPER"] = [SocialCalcFormula.InterestFunctions, -3, "nper", "", "financial"];
    FunctionClassList['financial']["PMT"] = [SocialCalcFormula.InterestFunctions, -3, "pmt", "", "financial"];
    FunctionClassList['financial']["PV"] = [SocialCalcFormula.InterestFunctions, -3, "pv", "", "financial"];
    FunctionClassList['financial']["RATE"] = [SocialCalcFormula.InterestFunctions, -3, "rate", "", "financial"];

    /*
    #
    # NPV(rate,v1,v2,c1:c2,...)
    #
    */

    SocialCalcFormula.NPVFunction = function (fname, operand, foperand) {

        var resulttypenpv, rate, sum, factor, value1;

        var scf = SocialCalcFormula;

        var rate = scf.OperandAsNumber(foperand);
        if (scf.CheckForErrorValue(operand, rate)) return;

        sum = 0;
        resulttypenpv = "n";
        factor = 1;

        while (foperand.length) {
            value1 = scf.OperandValueAndType(foperand);
            if (value1.type.charAt(0) == "n") {
                factor *= (1 + rate.value);
                if (factor == 0) {
                    scf.PushOperand(operand, "e#DIV/0!", 0);
                    return;
                }
                sum += value1.value / factor;
                resulttypenpv = scf.LookupResultType(value1.type, resulttypenpv || value1.type, scf.TypeLookupTable.plus);
            }
            else if (value1.type.charAt(0) == "e" && resulttypenpv.charAt(0) != "e") {
                resulttypenpv = value1.type;
                break;
            }
        }

        if (resulttypenpv.charAt(0) == "n") {
            resulttypenpv = 'n$';
        }

        scf.PushOperand(operand, resulttypenpv, sum);

        return;

    }

    FunctionClassList['financial']["NPV"] = [SocialCalcFormula.NPVFunction, -2, "npv", "", "financial"];

    /*
    #
    # IRR(c1:c2,[guess])
    #
    */

    SocialCalcFormula.IRRFunction = function (fname, operand, foperand) {

        var value1, guess, oldsum, maxloop, tries, epsilon, rate, oldrate, m, sum, factor, i;
        var rangeoperand = [];
        var cashflows = [];

        var scf = SocialCalcFormula;

        rangeoperand.push(foperand.pop()); // first operand is a range

        while (rangeoperand.length) { // get values from range so we can do iterative approximations
            value1 = scf.OperandValueAndType(rangeoperand);
            if (value1.type.charAt(0) == "n") {
                cashflows.push(value1.value);
            }
            else if (value1.type.charAt(0) == "e") {
                scf.PushOperand(operand, "e#VALUE!", 0);
                return;
            }
        }

        if (!cashflows.length) {
            scf.PushOperand(operand, "e#NUM!", 0);
            return;
        }

        guess = {value: 0};

        if (foperand.length) { // guess is provided
            guess = scf.OperandAsNumber(foperand);
            if (guess.type.charAt(0) != "n" && guess.type.charAt(0) != "b") {
                scf.PushOperand(operand, "e#VALUE!", 0);
                return;
            }
            if (foperand.length) { // should be no more args
                scf.FunctionArgsError(fname, operand);
                return;
            }
        }

        guess.value = guess.value || 0.1;

        // rate is calculated by repeated approximations
        // The deltas are used to calculate new guesses

        maxloop = 20;
        tries = 0;
        epsilon = 0.0000001; // this is close enough
        rate = guess.value;
        sum = 1;

        while ((sum >= 0 ? sum : -sum) > epsilon && (rate != oldrate)) {
            sum = 0;
            factor = 1;
            for (i = 0; i < cashflows.length; i++) {
                factor *= (1 + rate);
                if (factor == 0) {
                    scf.PushOperand(operand, "e#DIV/0!", 0);
                    return;
                }
                sum += cashflows[i] / factor;
            }

            if (oldsum != null) {
                m = (sum - oldsum) / (rate - oldrate); // get slope
                oldrate = rate;
                rate = rate - sum / m; // look for zero crossing
                oldsum = sum;
            }
            else { // first time - no old values
                oldrate = rate;
                rate = 1.1 * rate;
                oldsum = sum;
            }
            tries++;
            if (tries >= maxloop) { // didn't converge yet
                scf.PushOperand(operand, "e#NUM!", 0);
                return;
            }
        }

        scf.PushOperand(operand, 'n%', rate);

        return;

    }

    FunctionClassList['financial']["IRR"] = [SocialCalcFormula.IRRFunction, -1, "irr", "", "financial"];
    this.FunctionList = {}
    for (fclass in FunctionClassList) {
        for (fname in FunctionClassList[fclass]) {
            this.FunctionList[fname] = FunctionClassList[fclass][fname]
        }
    }

//
// SHEET CACHE
//

    // SocialCalcFormula.SheetCache = {
    //
    //     // Sheet data: Attributes are each sheet in the cache with values of an object with:
    //     //
    //     //    sheet: sheet-obj (or null, meaning not found)
    //     //    recalcstate: constants.asloaded = as loaded
    //     //                 constants.recalcing = being recalced now
    //     //                 constants.recalcdone = recalc done
    //     //    name: name of sheet (in case just have object and don't know name)
    //     //
    //
    //     sheets: {},
    //
    //     // Waiting for loading:
    //     // If sheet is not in cache, this is set to the sheetname being loaded
    //     // so it can be tested in the recalc loop to start load and then wait until restarted.
    //     // Reset to null before restarting.
    //
    //     waitingForLoading: null,
    //
    //     // Constants to use for setting sheets[*].recalcstate:
    //
    //     constants: {asloaded: 0, recalcing: 1, recalcdone: 2},
    //
    //     loadsheet: null // (deprecated - use SocialCalc.RecalcInfo.LoadSheet)
    //
    // };

//
// othersheet = SocialCalcFormula.FindInSheetCache(sheetname)
//
// Returns a SocialCalc.Sheet object corresponding to string sheetname
// or null if the sheet is not available or in error.
//
// Each sheet is loaded only once and then stored in a cache.
// Loading is handled elsewhere, e.g., in the recalc loop.
//

    // SocialCalcFormula.FindInSheetCache = function (sheetname) {
    //
    //     var str;
    //     var sfsc = SocialCalcFormula.SheetCache;
    //
    //     var nsheetname = SocialCalcFormula.NormalizeSheetName(sheetname); // normalize different versions
    //
    //     if (sfsc.sheets[nsheetname]) { // a sheet by that name is in the cache already
    //         return sfsc.sheets[nsheetname].sheet; // return it
    //     }
    //
    //     if (sfsc.waitingForLoading) { // waiting already - only queue up one
    //         return null; // return not found
    //     }
    //
    //     if (sfsc.loadsheet) { // Deprecated old format synchronous callback
    //         alert("Using SocialCalcFormula.SheetCache.loadsheet - deprecated");
    //         return SocialCalcFormula.AddSheetToCache(nsheetname, sfsc.loadsheet(nsheetname));
    //     }
    //
    //     sfsc.waitingForLoading = nsheetname; // let recalc loop know that we have a sheet to load
    //
    //     return null; // return not found
    //
    // }

//
// newsheet = SocialCalcFormula.AddSheetToCache(sheetname, str)
//
// Adds a new sheet to the sheet cache.
// Returns the sheet object filled out with the str (a saved sheet).
//

    // SocialCalcFormula.AddSheetToCache = function (sheetname, str) {
    //
    //     var newsheet = null;
    //     var sfsc = SocialCalcFormula.SheetCache;
    //     var sfscc = sfsc.constants;
    //     var newsheetname = SocialCalcFormula.NormalizeSheetName(sheetname);
    //
    //     if (str) {
    //         newsheet = new SocialCalc.Sheet();
    //         newsheet.ParseSheetSave(str);
    //     }
    //
    //     sfsc.sheets[newsheetname] = {sheet: newsheet, recalcstate: sfscc.asloaded, name: newsheetname};
    //
    //     SocialCalcFormula.FreshnessInfo.sheets[newsheetname] = true;
    //
    //     return newsheet;
    //
    // }

//
// nsheet = SocialCalcFormula.NormalizeSheetName(sheetname)
//

    // SocialCalcFormula.NormalizeSheetName = function (sheetname) {
    //
    //     if (SocialCalc.Callbacks.NormalizeSheetName) {
    //         return SocialCalc.Callbacks.NormalizeSheetName(sheetname);
    //     }
    //     else {
    //         return sheetname.toLowerCase();
    //     }
    // }

//
// REMOTE FUNCTION INFO
//

    // SocialCalcFormula.RemoteFunctionInfo = {
    //
    //     // Waiting for server:
    //     // If waiting for an XHR response from the server, this is set to some non-blank status text
    //     // so it can be tested in the recalc loop to start load and then wait until restarted.
    //     // Reset to null before restarting.
    //
    //     waitingForServer: null
    //
    // };

//
// FRESHNESS INFO
//
// This information is generated during recalc.
// It may be used to help determine when the recalc data in a spreadsheet
// may be out of date.
// For example, it may be used to display a message like:
// "Dependent on sheet 'FOO' which was updated more recently than this printout"

    // SocialCalcFormula.FreshnessInfo = {
    //
    //     // For each external sheet referenced successfully an attribute of that name with value true.
    //
    //     sheets: {},
    //
    //     // For each volatile function that is called an attribute of that name with value true.
    //
    //     volatile: {},
    //
    //     // Set to false when started and true when recalc completes
    //
    //     recalc_completed: false
    //
    // };
    //
    // SocialCalcFormula.FreshnessInfoReset = function () {
    //
    //     var scffi = SocialCalcFormula.FreshnessInfo;
    //
    //     scffi.sheets = {};
    //     scffi.volatile = {};
    //     scffi.recalc_completed = false;
    //
    // }

//
// MISC ROUTINES
//

//
// result = SocialCalcFormula.PlainCoord(coord)
//
// Returns: coord without any $'s
//

    // SocialCalcFormula.PlainCoord = function (coord) {
    //
    //     if (coord.indexOf("$") == -1) return coord;
    //
    //     return coord.replace(/\$/g, ""); // remove any $'s
    //
    // }

//
// result = SocialCalcFormula.OrderRangeParts(coord1, coord2)
//
// Returns: {c1: col, r1: row, c2: col, r2 = row} with c1/r1 upper left
//

    // SocialCalcFormula.OrderRangeParts = function (coord1, coord2) {
    //
    //     var cr1, cr2;
    //     var result = {};
    //
    //     cr1 = SocialCalc.coordToCr(coord1);
    //     cr2 = SocialCalc.coordToCr(coord2);
    //     if (cr1.col > cr2.col) {
    //         result.c1 = cr2.col;
    //         result.c2 = cr1.col;
    //     }
    //     else {
    //         result.c1 = cr1.col;
    //         result.c2 = cr2.col;
    //     }
    //     if (cr1.row > cr2.row) {
    //         result.r1 = cr2.row;
    //         result.r2 = cr1.row;
    //     }
    //     else {
    //         result.r1 = cr1.row;
    //         result.r2 = cr2.row;
    //     }
    //
    //     return result;
    //
    // }

//
// cond = SocialCalcFormula.TestCriteria(value, type, criteria)
//
// Determines whether a value/type meets the criteria.
// A criteria can be a numeric value, text beginning with <, <=, =, >=, >, <>, text by itself is start of text to match.
// Used by a variety of functions, including the "D" functions (DSUM, etc.).
//
// Returns true or false
//

    // SocialCalcFormula.TestCriteria = function (value, type, criteria) {
    //
    //     var comparitor, basestring, basevalue, cond, testvalue;
    //
    //     if (criteria == null) { // undefined (e.g., error value) is always false
    //         return false;
    //     }
    //
    //     criteria = criteria + "";
    //     comparitor = criteria.charAt(0); // look for comparitor
    //     if (comparitor == "=" || comparitor == "<" || comparitor == ">") {
    //         basestring = criteria.substring(1);
    //     }
    //     else {
    //         comparitor = criteria.substring(0, 2);
    //         if (comparitor == "<=" || comparitor == "<>" || comparitor == ">=") {
    //             basestring = criteria.substring(2);
    //         }
    //         else {
    //             comparitor = "none";
    //             basestring = criteria;
    //         }
    //     }
    //
    //     basevalue = SocialCalc.DetermineValueType(basestring); // get type of value being compared
    //     if (!basevalue.type) { // no criteria base value given
    //         if (comparitor == "none") { // blank criteria matches nothing
    //             return false;
    //         }
    //         if (type.charAt(0) == "b") { // comparing to empty cell
    //             if (comparitor == "=") { // empty equals empty
    //                 return true;
    //             }
    //         }
    //         else {
    //             if (comparitor == "<>") { // "something" does not equal empty
    //                 return true;
    //             }
    //         }
    //         return false; // otherwise false
    //     }
    //
    //     cond = false;
    //
    //     if (basevalue.type.charAt(0) == "n" && type.charAt(0) == "t") { // criteria is number, but value is text
    //         testvalue = SocialCalc.DetermineValueType(value);
    //         if (testvalue.type.charAt(0) == "n") { // could be number - make it one
    //             value = testvalue.value;
    //             type = testvalue.type;
    //         }
    //     }
    //
    //     if (type.charAt(0) == "n" && basevalue.type.charAt(0) == "n") { // compare two numbers
    //         value = value - 0; // make sure numbers
    //         basevalue.value = basevalue.value - 0;
    //         switch (comparitor) {
    //             case "<":
    //                 cond = value < basevalue.value;
    //                 break;
    //
    //             case "<=":
    //                 cond = value <= basevalue.value;
    //                 break;
    //
    //             case "=":
    //             case "none":
    //                 cond = value == basevalue.value;
    //                 break;
    //
    //             case ">=":
    //                 cond = value >= basevalue.value;
    //                 break;
    //
    //             case ">":
    //                 cond = value > basevalue.value;
    //                 break;
    //
    //             case "<>":
    //                 cond = value != basevalue.value;
    //                 break;
    //         }
    //     }
    //
    //     else if (type.charAt(0) == "e") { // error on left
    //         cond = false;
    //     }
    //
    //     else if (basevalue.type.charAt(0) == "e") { // error on right
    //         cond = false;
    //     }
    //
    //     else { // text, maybe mixed with number or blank
    //         if (type.charAt(0) == "n") {
    //             value = SocialCalc.format_number_for_display(value, "n", "");
    //         }
    //         if (basevalue.type.charAt(0) == "n") {
    //             return false; // if number and didn't match already, isn't a match
    //         }
    //
    //         value = value ? value.toLowerCase() : "";
    //         basevalue.value = basevalue.value ? basevalue.value.toLowerCase() : "";
    //
    //         switch (comparitor) {
    //             case "<":
    //                 cond = value < basevalue.value;
    //                 break;
    //
    //             case "<=":
    //                 cond = value <= basevalue.value;
    //                 break;
    //
    //             case "=":
    //                 cond = value == basevalue.value;
    //                 break;
    //
    //             case "none":
    //                 cond = value.substring(0, basevalue.value.length) == basevalue.value;
    //                 break;
    //
    //             case ">=":
    //                 cond = value >= basevalue.value;
    //                 break;
    //
    //             case ">":
    //                 cond = value > basevalue.value;
    //                 break;
    //
    //             case "<>":
    //                 cond = value != basevalue.value;
    //                 break;
    //         }
    //     }
    //
    //     return cond;
    //
    // }
}

Formula.prototype.ParseFormulaIntoTokens = function (line) {
    var i, ch, chclass, haddecimal, last_token, last_token_type, last_token_text, t;
    if(line == null) line = ''
    //var scf = SocialCalcFormula;
    var scc = this.scc
    var parsestate = this.ParseState;
    var tokentype = this.TokenType;
    var charclass = this.CharClass;
    var charclasstable = this.CharClassTable;
    var uppercasetable = this.UpperCaseTable; // much faster than toUpperCase function
    var pushtoken = this.ParsePushToken;
    var coordregex = /^\$?[A-Z]{1,2}\$?[1-9]\d*$/i;

    var parseinfo = [];
    var str = "";
    var state = 0;
    var haddecimal = false;

    for (i = 0; i <= line.length; i++) {
        if (i < line.length) {
            ch = line.charAt(i);
            cclass = charclasstable[ch];
        }
        else {
            ch = "";
            cclass = charclass.eof;
        }

        if (state == parsestate.num) {
            if (cclass == charclass.num) {
                str += ch;
            }
            else if (cclass == charclass.numstart && !haddecimal) {
                haddecimal = true;
                str += ch;
            }
            else if (ch == "E" || ch == "e") {
                str += ch;
                haddecimal = false;
                state = parsestate.numexp1;
            }
            else { // end of number - save it
                pushtoken(parseinfo, str, tokentype.num, 0);
                haddecimal = false;
                state = 0;
            }
        }

        if (state == parsestate.numexp1) {
            if (cclass == parsestate.num) {
                state = parsestate.numexp2;
            }
            else if ((ch == '+' || ch == '-') && (uppercasetable[str.charAt(str.length - 1)] == 'E')) {
                str += ch;
            }
            else if (ch == 'E' || ch == 'e') {
                ;
            }
            else {
                pushtoken(parseinfo, this.s_parseerrexponent, tokentype.error, 0);
                state = 0;
            }
        }

        if (state == parsestate.numexp2) {
            if (cclass == charclass.num) {
                str += ch;
            }
            else { // end of number - save it
                pushtoken(parseinfo, str, tokentype.num, 0);
                state = 0;
            }
        }

        if (state == parsestate.alpha) {
            if (cclass == charclass.num) {
                state = parsestate.coord;
            }
            else if (cclass == charclass.alpha || ch == ".") { // alpha may be letters, numbers, "_", or "."
                str += ch;
            }
            else if (cclass == charclass.incoord) {
                state = parsestate.coord;
            }
            else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
                pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
                state = 0;
            }
            else {
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
                state = 0;
            }
        }

        if (state == parsestate.coord) {
            if (cclass == charclass.num) {
                str += ch;
            }
            else if (cclass == charclass.incoord) {
                str += ch;
            }
            else if (cclass == charclass.alpha) {
                state = parsestate.alphanumeric;
            }
            else if (cclass == charclass.op || cclass == charclass.numstart ||
                cclass == charclass.eof || cclass == charclass.space) {
                if (coordregex.test(str)) {
                    t = tokentype.coord;
                }
                else {
                    t = tokentype.name;
                }
                pushtoken(parseinfo, str.toUpperCase(), t, 0);
                state = 0;
            }
            else {
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
                state = 0;
            }
        }


        if (state == parsestate.alphanumeric) {
            if (cclass == charclass.num || cclass == charclass.alpha) {
                str += ch;
            }
            else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
                pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
                state = 0;
            }
            else {
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
                state = 0;
            }
        }

        if (state == parsestate.string) {
            if (cclass == charclass.quote) {
                state = parsestate.stringquote; // got quote in string: is it doubled (quote in string) or by itself (end of string)?
            }
            else if (cclass == charclass.eof) {
                pushtoken(parseinfo, scc.s_parseerrstring, tokentype.error, 0);
                state = 0;
            }
            else {
                str += ch;
            }
        }
        else if (state == parsestate.stringquote) { // note else if here
            if (cclass == charclass.quote) {
                str += '"';
                state = parsestate.string; // double quote: add one then continue getting string
            }
            else { // something else -- end of string
                pushtoken(parseinfo, str, tokentype.string, 0);
                state = 0; // drop through to process
            }
        }

        else if (state == parsestate.specialvalue) { // special values like #REF!
            if (str.charAt(str.length - 1) == "!") { // done - save value as a name
                pushtoken(parseinfo, str, tokentype.name, 0);
                state = 0; // drop through to process
            }
            else if (cclass == charclass.eof) {
                pushtoken(parseinfo, scc.s_parseerrspecialvalue, tokentype.error, 0);
                state = 0;
            }
            else {
                str += ch;
            }
        }

        if (state == 0) {
            if (cclass == charclass.num) {
                str = ch;
                state = parsestate.num;
            }
            else if (cclass == charclass.numstart) {
                str = ch;
                haddecimal = true;
                state = parsestate.num;
            }
            else if (cclass == charclass.alpha || cclass == charclass.incoord) {
                str = ch;
                state = parsestate.alpha;
            }
            else if (cclass == charclass.specialstart) {
                str = ch;
                state = parsestate.specialvalue;
            }
            else if (cclass == charclass.op) {
                str = ch;
                if (parseinfo.length > 0) {
                    last_token = parseinfo[parseinfo.length - 1];
                    last_token_type = last_token.type;
                    last_token_text = last_token.text;
                    if (last_token_type == charclass.op) {
                        if (last_token_text == '<' || last_token_text == ">") {
                            str = last_token_text + str;
                            parseinfo.pop();
                            if (parseinfo.length > 0) {
                                last_token = parseinfo[parseinfo.length - 1];
                                last_token_type = last_token.type;
                                last_token_text = last_token.text;
                            }
                            else {
                                last_token_type = charclass.eof;
                                last_token_text = "EOF";
                            }
                        }
                    }
                }
                else {
                    last_token_type = charclass.eof;
                    last_token_text = "EOF";
                }
                t = tokentype.op;
                if ((parseinfo.length == 0)
                    || (last_token_type == charclass.op && last_token_text != ')' && last_token_text != '%')) { // Unary operator
                    if (str == '-') { // M is unary minus
                        str = "M";
                        ch = "M";
                    }
                    else if (str == '+') { // P is unary plus
                        str = "P";
                        ch = "P";
                    }
                    else if (str == ')' && last_token_text == '(') { // null arg list OK
                        ;
                    }
                    else if (str != '(') { // binary-op open-paren OK, others no
                        t = tokentype.error;
                        str = this.s_parseerrtwoops;
                    }
                }
                else if (str.length > 1) {
                    if (str == '>=') { // G is >=
                        str = "G";
                        ch = "G";
                    }
                    else if (str == '<=') { // L is <=
                        str = "L";
                        ch = "L";
                    }
                    else if (str == '<>') { // N is <>
                        str = "N";
                        ch = "N";
                    }
                    else {
                        t = tokentype.error;
                        str = scc.s_parseerrtwoops;
                    }
                }
                pushtoken(parseinfo, str, t, ch);
                state = 0;
            }
            else if (cclass == charclass.quote) { // starting a string
                str = "";
                state = parsestate.string;
            }
            else if (cclass == charclass.space) { // store so can reconstruct spacing
                pushtoken(parseinfo, " ", tokentype.space, 0);
            }
            else if (cclass == charclass.eof) { // ignore -- needed to have extra loop to close out other things
            }
            else { // unknown class - such as unknown char
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            }
        }
    }
    //console.log(parseinfo)
    return parseinfo;

}
Formula.prototype.ParsePushToken = function (parseinfo, ttext, ttype, topcode) {

    parseinfo.push({text: ttext, type: ttype, opcode: topcode});

}

Formula.prototype.evaluate_parsed_formula = function (parseinfo, allowrangereturn) {

    var result;

    var scf = this
    var tokentype = scf.TokenType;

    var revpolish;
    var parsestack = [];

    var errortext = "";
    //console.log(parseinfo)
    revpolish = this.ConvertInfixToPolish(parseinfo); // result is either an array or a string with error text
    //console.log(revpolish)
    result = this.EvaluatePolish(parseinfo, revpolish, allowrangereturn);
    //console.log(result)
    return result;

}
Formula.prototype.ConvertInfixToPolish = function (parseinfo) {

    var scf = this
    var scc = this.Constants;
    var tokentype = scf.TokenType;
    var token_precedence = scf.TokenPrecedence;

    var revpolish = [];
    var parsestack = [];

    var errortext = "";

    var function_start = -1;

    var i, pii, ttype, ttext, tprecedence, tstackprecedence;

    for (i = 0; i < parseinfo.length; i++) {
        pii = parseinfo[i];
        ttype = pii.type;
        ttext = pii.text;
        if (ttype == tokentype.num || ttype == tokentype.coord || ttype == tokentype.string) {
            revpolish.push(i);
        }
        else if (ttype == tokentype.name) {
            parsestack.push(i);
            revpolish.push(function_start);
        }
        else if (ttype == tokentype.space) { // ignore
            continue;
        }
        else if (ttext == ',') {
            while (parsestack.length && parseinfo[parsestack[parsestack.length - 1]].text != "(") {
                revpolish.push(parsestack.pop());
            }
            if (parsestack.length == 0) { // no ( -- error
                errortext = scc.s_parseerrmissingopenparen;
                break;
            }
        }
        else if (ttext == '(') {
            parsestack.push(i);
        }
        else if (ttext == ')') {
            while (parsestack.length && parseinfo[parsestack[parsestack.length - 1]].text != "(") {
                revpolish.push(parsestack.pop());
            }
            if (parsestack.length == 0) { // no ( -- error
                errortext = scc.s_parseerrcloseparennoopen;
                break;
            }
            parsestack.pop();
            if (parsestack.length && parseinfo[parsestack[parsestack.length - 1]].type == tokentype.name) {
                revpolish.push(parsestack.pop());
            }
        }
        else if (ttype == tokentype.op) {
            if (parsestack.length && parseinfo[parsestack[parsestack.length - 1]].type == tokentype.name) {
                revpolish.push(parsestack.pop());
            }
            while (parsestack.length && parseinfo[parsestack[parsestack.length - 1]].type == tokentype.op
            && parseinfo[parsestack[parsestack.length - 1]].text != '(') {
                tprecedence = token_precedence[pii.opcode];
                tstackprecedence = token_precedence[parseinfo[parsestack[parsestack.length - 1]].opcode];
                if (tprecedence >= 0 && tprecedence < tstackprecedence) {
                    break;
                }
                else if (tprecedence < 0) {
                    tprecedence = -tprecedence;
                    if (tstackprecedence < 0) tstackprecedence = -tstackprecedence;
                    if (tprecedence <= tstackprecedence) {
                        break;
                    }
                }
                revpolish.push(parsestack.pop());
            }
            parsestack.push(i);
        }
        else if (ttype == tokentype.error) {
            errortext = ttext;
            break;
        }
        else {
            errortext = "Internal error while processing parsed formula. ";
            break;
        }
    }
    while (parsestack.length > 0) {
        if (parseinfo[parsestack[parsestack.length - 1]].text == '(') {
            errortext = scc.s_parseerrmissingcloseparen;
            break;
        }
        revpolish.push(parsestack.pop());
    }

    if (errortext) {
        return errortext;
    }

    return revpolish;

}
Formula.prototype.EvaluatePolish = function (parseinfo, revpolish, allowrangereturn) {

    var scf = this
    var scc = this.scc
    var tokentype = this.TokenType
    var lookup_result_type = this.LookupResultType
    var typelookup = this.TypeLookupTable

    //var operand_as_text = this.OperandAsText
    //var operand_value_and_type = this.OperandValueAndType
    var operands_as_coord_on_sheet = this.OperandsAsCoordOnSheet
    //var format_number_for_display = SocialCalc.format_number_for_display || function(v, t, f) {return v+"";};

    var errortext = "";
    var function_start = -1;
    var missingOperandError = {value: "", type: "e#VALUE!", error: scc.s_parseerrmissingoperand};

    var operand = [];
    var PushOperand = function (t, v) {
        operand.push({type: t, value: v});
    };

    var i, rii, prii, ttype, ttext, value1, value2, tostype, tostype2, resulttype, valuetype, cond, vmatch, smatch;

    if (!parseinfo.length || (!(revpolish instanceof Array))) {
        return ({value: "", type: "e#VALUE!", error: (typeof revpolish == "string" ? revpolish : "")});
    }

    for (i = 0; i < revpolish.length; i++) {
        rii = revpolish[i];
        if (rii == function_start) { // Remember the start of a function argument list
            PushOperand("start", 0);
            continue;
        }

        prii = parseinfo[rii];
        ttype = prii.type;
        ttext = prii.text;
        //console.log(prii)
        if (ttype == tokentype.num) {
            PushOperand("n", ttext - 0);
        }

        else if (ttype == tokentype.coord) {
            PushOperand("coord", ttext);
        }

        else if (ttype == tokentype.string) {
            PushOperand("t", ttext);
        }

        else if (ttype == tokentype.op) {
            if (operand.length <= 0) { // Nothing on the stack...
                return missingOperandError;
                break; // done
            }

            // Unary minus

            if (ttext == 'M') {
                value1 = this.OperandAsNumber(operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.unaryminus);
                PushOperand(resulttype, -value1.value);
            }

            // Unary plus

            else if (ttext == 'P') {
                value1 = this.OperandAsNumber(operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.unaryplus);
                PushOperand(resulttype, value1.value);
            }

            // Unary % - percent, left associative

            else if (ttext == '%') {
                value1 = this.OperandAsNumber(operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.unarypercent);
                PushOperand(resulttype, 0.01 * value1.value);
            }

            // & - string concatenate

            else if (ttext == '&') {
                if (operand.length <= 1) { // Need at least two things on the stack...
                    return missingOperandError;
                }
                value2 = this.OperandAsText(operand);
                value1 = this.OperandAsText(operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.concat);
                PushOperand(resulttype, value1.value + value2.value);
            }

            // : - Range constructor

            else if (ttext == ':') {
                if (operand.length <= 1) { // Need at least two things on the stack...
                    return missingOperandError;
                }
                value1 = scf.OperandsAsRangeOnSheet(operand); // get coords even if use name on other sheet
                if (value1.error) { // not available
                    errortext = errortext || value1.error;
                }
                //console.log(value1)
                PushOperand(value1.type, value1.value); // push sheetname with range on that sheet
            }

            // ! - sheetname!coord

            else if (ttext == '!') {
                if (operand.length <= 1) { // Need at least two things on the stack...
                    return missingOperandError;
                }
                value1 = operands_as_coord_on_sheet(operand); // get coord even if name on other sheet
                if (value1.error) { // not available
                    errortext = errortext || value1.error;
                }
                PushOperand(value1.type, value1.value); // push sheetname with coord or range on that sheet
            }

            // Comparison operators: < L = G > N (< <= = >= > <>)

            else if (ttext == "<" || ttext == "L" || ttext == "=" || ttext == "G" || ttext == ">" || ttext == "N") {
                if (operand.length <= 1) { // Need at least two things on the stack...
                    errortext = scc.s_parseerrmissingoperand; // remember error
                    break;
                }
                value2 = this.OperandValueAndType(operand);
                value1 = this.OperandValueAndType(operand);
                if (value1.type.charAt(0) == "n" && value2.type.charAt(0) == "n") { // compare two numbers
                    cond = 0;
                    if (ttext == "<") {
                        cond = value1.value < value2.value ? 1 : 0;
                    }
                    else if (ttext == "L") {
                        cond = value1.value <= value2.value ? 1 : 0;
                    }
                    else if (ttext == "=") {
                        cond = value1.value == value2.value ? 1 : 0;
                    }
                    else if (ttext == "G") {
                        cond = value1.value >= value2.value ? 1 : 0;
                    }
                    else if (ttext == ">") {
                        cond = value1.value > value2.value ? 1 : 0;
                    }
                    else if (ttext == "N") {
                        cond = value1.value != value2.value ? 1 : 0;
                    }
                    PushOperand("nl", cond);
                }
                else if (value1.type.charAt(0) == "e") { // error on left
                    PushOperand(value1.type, 0);
                }
                else if (value2.type.charAt(0) == "e") { // error on right
                    PushOperand(value2.type, 0);
                }
                else { // text maybe mixed with numbers or blank
                    tostype = value1.type.charAt(0);
                    tostype2 = value2.type.charAt(0);
                    if (tostype == "n") {
                        //value1.value = format_number_for_display(value1.value, "n", "");
                    }
                    else if (tostype == "b") {
                        value1.value = "";
                    }
                    if (tostype2 == "n") {
                        //value2.value = //format_number_for_display(value2.value, "n", "");
                    }
                    else if (tostype2 == "b") {
                        value2.value = "";
                    }
                    cond = 0;
                    value1.value = value1.value.toLowerCase(); // ignore case
                    value2.value = value2.value.toLowerCase();
                    if (ttext == "<") {
                        cond = value1.value < value2.value ? 1 : 0;
                    }
                    else if (ttext == "L") {
                        cond = value1.value <= value2.value ? 1 : 0;
                    }
                    else if (ttext == "=") {
                        cond = value1.value == value2.value ? 1 : 0;
                    }
                    else if (ttext == "G") {
                        cond = value1.value >= value2.value ? 1 : 0;
                    }
                    else if (ttext == ">") {
                        cond = value1.value > value2.value ? 1 : 0;
                    }
                    else if (ttext == "N") {
                        cond = value1.value != value2.value ? 1 : 0;
                    }
                    PushOperand("nl", cond);
                }
            }

            // Normal infix arithmethic operators: +, -. *, /, ^

            else { // what's left are the normal infix arithmetic operators
                if (operand.length <= 1) { // Need at least two things on the stack...
                    errortext = scc.s_parseerrmissingoperand; // remember error
                    break;
                }
                value2 = this.OperandAsNumber(operand);
                value1 = this.OperandAsNumber(operand);
                if (ttext == '+') {
                    resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
                    PushOperand(resulttype, value1.value + value2.value);
                }
                else if (ttext == '-') {
                    resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
                    PushOperand(resulttype, value1.value - value2.value);
                }
                else if (ttext == '*') {
                    resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
                    PushOperand(resulttype, value1.value * value2.value);
                }
                else if (ttext == '/') {
                    if (value2.value != 0) {
                        PushOperand("n", value1.value / value2.value); // gives plain numeric result type
                    }
                    else {
                        PushOperand("e#DIV/0!", 0);
                    }
                }
                else if (ttext == '^') {
                    value1.value = Math.pow(value1.value, value2.value);
                    value1.type = "n"; // gives plain numeric result type
                    if (isNaN(value1.value)) {
                        value1.value = 0;
                        value1.type = "e#NUM!";
                    }
                    PushOperand(value1.type, value1.value);
                }
            }
        }

        // function or name

        else if (ttype == tokentype.name) {
            errortext = scf.CalculateFunction(ttext, operand);
            if (errortext) break;
        }

        else {
            errortext = scc.s_InternalError + "Unknown token " + ttype + " (" + ttext + "). ";
            break;
        }
    }
    // look at final value and handle special cases

    value = operand[0] ? operand[0].value : "";
    tostype = operand[0] ? operand[0].type : "";

    // if (tostype == "name") { // name - expand it
    //     value1 = SocialCalcFormula.LookupName(sheet, value);
    //     value = value1.value;
    //     tostype = value1.type;
    //     errortext = errortext || value1.error;
    // }

    if (tostype == "coord") { // the value is a coord reference, get its value and type
        value1 = this.OperandValueAndType(operand);
        value = value1.value;
        tostype = value1.type;
        if (tostype == "b") {
            tostype = "n";
            value = 0;
        }
    }

    if (operand.length > 1 && !errortext) { // something left - error
        errortext += scc.s_parseerrerrorinformula;
    }

    // set return type

    valuetype = tostype;

    if (tostype.charAt(0) == "e") { // error value
        errortext = errortext || tostype.substring(1) || scc.s_calcerrerrorvalueinformula;
    }
    // else if (tostype == "range") {
    //     vmatch = value.match(/^(.*)\|(.*)\|/);
    //     smatch = vmatch[1].indexOf("!");
    //     if (smatch>=0) { // swap sheetname
    //         vmatch[1] = vmatch[1].substring(smatch+1) + "!" + vmatch[1].substring(0, smatch).toUpperCase();
    //     }
    //     else {
    //         vmatch[1] = vmatch[1].toUpperCase();
    //     }
    //     value = vmatch[1] + ":" + vmatch[2].toUpperCase();
    //     if (!allowrangereturn) {
    //         errortext = scc.s_formularangeresult+" "+value;
    //     }
    // }

    if (errortext && valuetype.charAt(0) != "e") {
        value = errortext;
        valuetype = "e";
    }

    // look for overflow

    if (valuetype.charAt(0) == "n" && (isNaN(value) || !isFinite(value))) {
        value = 0;
        valuetype = "e#NUM!";
        errortext = isNaN(value) ? scc.s_calcerrnumericnan : scc.s_calcerrnumericoverflow;
    }

    return ({value: value, type: valuetype, error: errortext});

}

Formula.prototype.LookupResultType = function (type1, type2, typelookup) {

    var pos1, pos2, result;

    var table1 = typelookup[type1];

    if (!table1) {
        table1 = typelookup[type1.charAt(0) + '*'];
        if (!table1) {
            return "e#VALUE! (internal error, missing LookupResultType " + type1.charAt(0) + "*)"; // missing from table -- please add it
        }
    }
    pos1 = table1.indexOf("|" + type2 + ":");
    if (pos1 >= 0) {
        pos2 = table1.indexOf("|", pos1 + 1);
        if (pos2 < 0) return "e#VALUE! (internal error, incorrect LookupResultType " + table1 + ")";
        result = table1.substring(pos1 + type2.length + 2, pos2);
        if (result == "1") return type1;
        if (result == "2") return type2;
        return result;
    }
    pos1 = table1.indexOf("|" + type2.charAt(0) + "*:");
    if (pos1 >= 0) {
        pos2 = table1.indexOf("|", pos1 + 1);
        if (pos2 < 0) return "e#VALUE! (internal error, incorrect LookupResultType " + table1 + ")";
        result = table1.substring(pos1 + 4, pos2);
        if (result == "1") return type1;
        if (result == "2") return type2;
        return result;
    }
    return "e#VALUE!";

}

Formula.prototype.OperandAsNumber = function (operand) {

    var t, valueinfo;
    var operandinfo = this.OperandValueAndType(operand)
    //console.log(operandinfo)
    t = operandinfo.type.charAt(0);

    if (t == "n") {
        operandinfo.value = operandinfo.value - 0;
    }
    else if (t == "b") { // blank cell
        operandinfo.type = "n";
        operandinfo.value = 0;
    }
    else if (t == "e") { // error
        operandinfo.value = 0;
    }
    else {
        valueinfo = this.DetermineValueType ? this.DetermineValueType(operandinfo.value) :
            {value: operandinfo.value - 0, type: "n"}; // if without rest of SocialCalc
        if (valueinfo.type.charAt(0) == "n") {
            operandinfo.value = valueinfo.value - 0;
            operandinfo.type = valueinfo.type;
        }
        else {
            operandinfo.value = 0;
            operandinfo.type = valueinfo.type;
        }
    }

    return operandinfo;

}
Formula.prototype.DetermineValueType = function (rawvalue) {

    var value = rawvalue + "";
    var type = "t";
    var tvalue, matches, year, hour, minute, second, denom, num, intgr, constr;

    tvalue = value.replace(/^\s+/, ""); // remove leading and trailing blanks
    tvalue = tvalue.replace(/\s+$/, "");

    if (value.length == 0) {
        type = "";
    }
    else if (value.match(/^\s+$/)) { // just blanks
        ; // leave type "t"
    }
    else if (tvalue.match(/^[-+]?\d*(?:\.)?\d*(?:[eE][-+]?\d+)?$/)) { // general number, including E
        value = tvalue - 0; // try converting to number
        if (isNaN(value)) { // leave alone - catches things like plain "-"
            value = rawvalue + "";
        }
        else {
            type = "n";
        }
    }
    else if (tvalue.match(/^[-+]?\d*(?:\.)?\d*\s*%$/)) { // percent form: 15.1%
        value = (tvalue.slice(0, -1) - 0) / 100; // convert and scale
        type = "n%";
    }
    else if (tvalue.match(/^[-+]?\$\s*\d*(?:\.)?\d*\s*$/) && tvalue.match(/\d/)) { // $ format: $1.49
        value = tvalue.replace(/\$/, "") - 0;
        type = "n$";
    }
    else if (tvalue.match(/^[-+]?(\d*,\d*)+(?:\.)?\d*$/)) { // number format ignoring commas: 1,234.49
        value = tvalue.replace(/,/g, "") - 0;
        type = "n";
    }
    else if (tvalue.match(/^[-+]?(\d*,\d*)+(?:\.)?\d*\s*%$/)) { // % with commas: 1,234.49%
        value = (tvalue.replace(/[%,]/g, "") - 0) / 100;
        type = "n%";
    }
    else if (tvalue.match(/^[-+]?\$\s*(\d*,\d*)+(?:\.)?\d*$/) && tvalue.match(/\d/)) { // $ and commas: $1,234.49
        value = tvalue.replace(/[\$,]/g, "") - 0;
        type = "n$";
    }
    else if (matches = value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{1,4})\s*$/)) { // MM/DD/YYYY, MM/DD/YYYY
        year = matches[3] - 0;
        year = year < 1000 ? year + 2000 : year;
        value = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, matches[1] - 0, matches[2] - 0) - 2415019;
        type = "nd";
    }
    else if (matches = value.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s*$/)) { // YYYY-MM-DD, YYYY/MM/DD
        year = matches[1] - 0;
        year = year < 1000 ? year + 2000 : year;
        value = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, matches[2] - 0, matches[3] - 0) - 2415019;
        type = "nd";
    }
    else if (matches = value.match(/^(\d{1,2}):(\d{1,2})\s*$/)) { // HH:MM
        hour = matches[1] - 0;
        minute = matches[2] - 0;
        if (hour < 24 && minute < 60) {
            value = hour / 24 + minute / (24 * 60);
            type = "nt";
        }
    }
    else if (matches = value.match(/^(\d{1,2}):(\d{1,2}):(\d{1,2})\s*$/)) { // HH:MM:SS
        hour = matches[1] - 0;
        minute = matches[2] - 0;
        second = matches[3] - 0;
        if (hour < 24 && minute < 60 && second < 60) {
            value = hour / 24 + minute / (24 * 60) + second / (24 * 60 * 60);
            type = "nt";
        }
    }
    else if (matches = value.match(/^\s*([-+]?\d+) (\d+)\/(\d+)\s*$/)) { // 1 1/2
        intgr = matches[1] - 0;
        num = matches[2] - 0;
        denom = matches[3] - 0;
        if (denom && denom > 0) {
            value = intgr + (intgr < 0 ? -num / denom : num / denom);
            type = "n";
        }
    }
    // else if (constr = SocialCalc.InputConstants[value.toUpperCase()]) { // special constants, like "false" and #N/A
    //     num = constr.indexOf(",");
    //     value = constr.substring(0, num) - 0;
    //     type = constr.substring(num + 1);
    // }

    else if (tvalue.length > 7 && tvalue.substring(0, 7).toLowerCase() == "http://") { // URL
        value = tvalue;
        type = "tl";
    }

    return {value: value, type: type};

}
Formula.prototype.OperandAsText = function (operand) {

    var t, valueinfo;
    var operandinfo = this.OperandValueAndType(operand);

    t = operandinfo.type.charAt(0);

    if (t == "t") { // any flavor of text returns as is
        ;
    }
    // else if (t == "n") {
    //     operandinfo.value = SocialCalc.format_number_for_display ?
    //         SocialCalc.format_number_for_display(operandinfo.value, operandinfo.type, "") :
    //         operandinfo.value = operandinfo.value+"";
    //     operandinfo.type = "t";
    // }
    else if (t == "b") { // blank
        operandinfo.value = "";
        operandinfo.type = "t";
    }
    else if (t == "e") { // error
        operandinfo.value = "";
    }
    else {
        operand.value = operandinfo.value + "";
        operand.type = "t";
    }

    return operandinfo;

}
Formula.prototype.OperandsAsCoordOnSheet = function (operand) {

    var sheetname, othersheet, pos1, pos2;
    var value1 = {};
    var result = {};
    var scf = this

    var stacklen = operand.length;
    value1.value = operand[stacklen - 1].value; // get top of stack - coord or name
    value1.type = operand[stacklen - 1].type;
    operand.pop(); // we have data - pop stack

    sheetname = scf.OperandAsSheetName(operand); // get sheetname as text
    othersheet = scf.FindInSheetCache(sheetname.value);
    if (othersheet == null) { // unavailable
        result.type = "e#REF!";
        result.value = 0;
        result.error = this.scc.s_sheetunavailable + " " + sheetname.value;
        return result;
    }

    if (value1.type == "name") {
        value1 = scf.LookupName(othersheet, value1.value);
    }
    result.type = value1.type;
    if (value1.type == "coord") { // value is a coord reference
        result.value = value1.value + "!" + sheetname.value; // return in the format as used on stack
    }
    else if (value1.type == "range") { // value is a range reference
        pos1 = value1.value.indexOf("|");
        pos2 = value1.value.indexOf("|", pos1 + 1);
        result.value = value1.value.substring(0, pos1) + "!" + sheetname.value +
            "|" + value1.value.substring(pos1 + 1, pos2) + "|";
    }
    else if (value1.type.charAt(0) == "e") {
        result.value = value1.value;
    }
    else {
        result.error = this.scc.s_calcerrcellrefmissing;
        result.type = "e#REF!";
        result.value = 0;
    }
    //console.log(result)
    return result;

}
Formula.prototype.OperandsAsRangeOnSheet = function (operand) {

    var value1, pos1, pos2;
    var value2 = {};
    var scf = this
    var scc = this.scc

    var stacklen = operand.length;
    value2.value = operand[stacklen - 1].value; // get top of stack - coord or name for "right" side
    value2.type = operand[stacklen - 1].type;
    operand.pop(); // we have data - pop stack

    value1 = scf.OperandAsCoord(operand); // get "left" coord
    if (value1.type != "coord") { // not a coord, which it must be
        return {value: 0, type: "e#REF!"};
    }

    pos1 = value1.value.indexOf("!");
    if (pos1 != -1) { // sheet reference
        pos2 = value1.value.indexOf("|", pos1 + 1);
        if (pos2 < 0) pos2 = value1.value.length;
        othersheet = scf.FindInSheetCache(value1.value.substring(pos1 + 1, pos2)); // get other sheet
        if (othersheet == null) { // unavailable
            return {
                value: 0,
                type: "e#REF!",
                errortext: scc.s_sheetunavailable + " " + value1.value.substring(pos1 + 1, pos2)
            };
        }
    }

    if (value2.type == "name") { // coord:name is allowed, if name is just one cell
        value2 = scf.LookupName(othersheet, value2.value);
    }

    if (value2.type == "coord") { // value is a coord reference, so return the combined range
        return {value: value1.value + "|" + value2.value + "|", type: "range"}; // return range in the format as used on stack
    }
    else { // bad form
        return {value: scc.s_calcerrcellrefmissing, type: "e#REF!"};
    }
}
Formula.prototype.OperandAsCoord = function (operand) {

    var scf = this

    var result = {type: "", value: ""};

    var stacklen = operand.length;

    result.value = operand[stacklen - 1].value; // get top of stack
    result.type = operand[stacklen - 1].type;
    operand.pop(); // we have data - pop stack
    // if (result.type == "name") {
    //     result = SocialCalc.Formula.LookupName(sheet, result.value);
    // }
    if (result.type == "coord") { // value is a coord reference
        return result;
    }
    else {
        result.value = this.scc.s_calcerrcellrefmissing;
        result.type = "e#REF!";
        return result;
    }
}
Formula.prototype.OperandValueAndType = function (operand) {
    var cellvtype, cell, pos, coordsheet;
    var scf = this;

    var result = {type: "", value: ""};

    var stacklen = operand.length;

    if (!stacklen) { // make sure something is there
        //result.error = SocialCalc.Constants.s_InternalError+"no operand on stack";
        return result;
    }
    var sw = false
    result.value = operand[stacklen - 1].value; // get top of stack
    result.type = operand[stacklen - 1].type;
    operand.pop(); // we have data - pop stack

    if (result.type == "name") {
        //result = scf.LookupName( result.value);
        //console.log('Name!')
    }

    if (result.type == "range") {
        //result = scf.StepThroughRangeDown(operand, result.value);
        //console.log('Range!')
    }

    if (result.type == "coord") { // value is a coord reference
        var pcoord1 = result.value.match(/[A-Za_z]+/g)
        var pcoord2 = result.value.match(/\d+/g)
        var pid = pcoord1[0].toUpperCase()+'_'+pcoord2[0]
        if(!this.sheet.namecc){
            this.sheet.namecc = {}
            this.sheet.namecc[this.sheet.fid] = true
            sw = true
        }
        if(this.sheet.namecc[pid]){
            result = {value: 'erro', type: 'e', error: 'erro'}
        }
        else{
            this.sheet.namecc[pid] = true
            //console.log(this.sheet.namecc)
            if(this.sheet.cells[pid]){
                if(this.sheet.cells[pid].formula !== ''){
                    var parseinfo = this.ParseFormulaIntoTokens(this.sheet.cells[pid].formula.substring(1))
                    result = this.evaluate_parsed_formula(parseinfo, 1)
                }
                else{
                    var parseinfo = this.ParseFormulaIntoTokens(this.sheet.cells[pid].content)
                    if(parseinfo.length === 1) result = this.evaluate_parsed_formula(parseinfo, 1)
                    else result = {value: this.sheet.cells[pid].content+'', type: 's', error: ''}
                }
            }
            else result = {value: 0, type: 'n', error: ''}
        }



    }
    if(sw) delete this.sheet.namecc
    return result;

}
Formula.prototype.CalculateFunction = function (fname, operand) {
    var fobj, foperand, ffunc, argnum, ttext;
    var scf = this
    var ok = 1;
    var errortext = "";

    fobj = scf.FunctionList[fname];
    //console.log(fobj)
    if (fobj) {
        foperand = [];
        ffunc = fobj[0];
        argnum = fobj[1];
        scf.CopyFunctionArgs(operand, foperand);
        if (argnum != 100) {
            if (argnum < 0) {
                if (foperand.length < -argnum) {
                    errortext = scf.FunctionArgsError(fname, operand);
                    return errortext;
                }
            }
            else {
                //console.log(foperand)
                if (foperand.length != argnum) {
                    errortext = scf.FunctionArgsError(fname, operand);
                    return errortext;
                }
            }
        }
        errortext = ffunc(fname, operand, foperand);
    }

    else {
        ttext = fname;

        if (operand.length && operand[operand.length - 1].type == "start") { // no arguments - name or zero arg function
            operand.pop();
            scf.PushOperand(operand, "name", ttext);
        }

        else {
            errortext = SocialCalc.Constants.s_sheetfuncunknownfunction + " " + ttext + ". ";
        }
    }

    return errortext;

}
Formula.prototype.CopyFunctionArgs = function (operand, foperand) {

    var fobj, ffunc, argnum;
    var scf = this
    var ok = 1;
    var errortext = null;

    while (operand.length > 0 && operand[operand.length - 1].type != "start") { // get each arg
        foperand.push(operand.pop()); // copy it
    }
    operand.pop(); // get rid of "start"

    return;

}
Formula.prototype.FunctionArgsError = function (fname, operand) {

    var errortext = this.scc.s_calcerrincorrectargstofunction + " " + fname + ". ";
    this.PushOperand(operand, "e#VALUE!", errortext);

    return errortext;

}
Formula.prototype.DecodeRangeParts = function (sheetdata, range) {

    var value1, value2, pos1, pos2, sheet1, coordsheetdata, rp;

    var scf = SocialCalc.Formula;

    pos1 = range.indexOf("|");
    pos2 = range.indexOf("|", pos1 + 1);
    value1 = range.substring(0, pos1);
    value2 = range.substring(pos1 + 1, pos2);

    pos1 = value1.indexOf("!");
    if (pos1 != -1) {
        sheet1 = value1.substring(pos1 + 1);
        value1 = value1.substring(0, pos1);
    }
    else {
        sheet1 = "";
    }
    pos1 = value2.indexOf("!");
    if (pos1 != -1) {
        value2 = value2.substring(0, pos1);
    }

    coordsheetdata = sheetdata;
    if (sheet1) { // sheet reference
        coordsheetdata = scf.FindInSheetCache(sheet1);
        if (coordsheetdata == null) { // unavailable
            return null;
        }
    }

    rp = scf.OrderRangeParts(value1, value2);

    return {
        sheetdata: coordsheetdata,
        sheetname: sheet1,
        col1num: rp.c1,
        ncols: rp.c2 - rp.c1 + 1,
        row1num: rp.r1,
        nrows: rp.r2 - rp.r1 + 1
    }

}

module.exports.Formula = Formula