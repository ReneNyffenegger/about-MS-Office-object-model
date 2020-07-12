option explicit

dim r as long

sub main() ' {


   P  xl24HourClock            , "xl24HourClock"             ' 33  True if you are using 24-hour time; False if you are using 12-hour time.
   P  xl4DigitYears            , "xl4DigitYears"             ' 43  True if you are using four-digit years; False if you are using two-digit years.
   P  xlAlternateArraySeparator, "xlAlternateArraySeparator" ' 16  Alternate array item separator to be used if the current array separator is the same as the decimal separator.
   P  xlColumnSeparator        , "xlColumnSeparator"         ' 14  Character used to separate columns in array literals.
   P  xlCountryCode            , "xlCountryCode"             '  1  Country/Region version of Microsoft Excel.
   P  xlCountrySetting         , "xlCountrySetting"          '  2  Current country/region setting in the Windows Control Panel.
   P  xlCurrencyBefore         , "xlCurrencyBefore"          ' 37  True if the currency symbol precedes the currency values; False if it follows them.
   P  xlCurrencyCode           , "xlCurrencyCode"            ' 25  Currency symbol.
   P  xlCurrencyDigits         , "xlCurrencyDigits"          ' 27  Number of decimal digits to be used in currency formats.
   P  xlCurrencyLeadingZeros   , "xlCurrencyLeadingZeros"    ' 40  True if leading zeros are displayed for zero currency values.
   P  xlCurrencyMinusSign      , "xlCurrencyMinusSign"       ' 38  True if you are using a minus sign for negative numbers; False if you are using parentheses.
   P  xlCurrencyNegative       , "xlCurrencyNegative"        ' 28  Currency format for negative currency values: 0 = (symbolx) or (xsymbol), 1 = -symbolx or -xsymbol, 2 = symbol-x or x-symbol, or 3 = symbolx- or xsymbol-, where symbol is the currency symbol of the country or region.  Note that the position of the currency symbol is determined by xlCurrencyBefore.
   P  xlCurrencySpaceBefore    , "xlCurrencySpaceBefore"     ' 36  True if a space is added before the currency symbol.
   P  xlCurrencyTrailingZeros  , "xlCurrencyTrailingZeros"   ' 39  True if trailing zeros are displayed for zero currency values.
   P  xlDateOrder              , "xlDateOrder"               ' 32  Order of date elements: 0 = month-day-year, 1 = day-month-year, 2 = year-month-day
   P  xlDateSeparator          , "xlDateSeparator"           ' 17  Date separator (/).
   P  xlDayCode                , "xlDayCode"                 ' 21  Day symbol (d).
   P  xlDayLeadingZero         , "xlDayLeadingZero"          ' 42  True if a leading zero is displayed in days.
   P  xlDecimalSeparator       , "xlDecimalSeparator"        '  3  Decimal separator.
   P  xlGeneralFormatName      , "xlGeneralFormatName"       ' 26  Name of the General number format.
   P  xlHourCode               , "xlHourCode"                ' 22  Hour symbol (h).
   P  xlLeftBrace              , "xlLeftBrace"               ' 12  Character used instead of the left brace ({) in array literals.
   P  xlLeftBracket            , "xlLeftBracket"             ' 10  Character used instead of the left bracket ([) in R1C1-style relative references.
   P  xlListSeparator          , "xlListSeparator"           '  5  List separator.
   P  xlLowerCaseColumnLetter  , "xlLowerCaseColumnLetter"   '  9  Lowercase column letter.
   P  xlLowerCaseRowLetter     , "xlLowerCaseRowLetter"      '  8  Lowercase row letter.
   P  xlMDY                    , "xlMDY"                     ' 44  True if the date order is month-day-year for dates displayed in the long form; False if the date order is day-month-year.
   P  xlMetric                 , "xlMetric"                  ' 35  True if you are using the metric system; False if you are using the English measurement system.
   P  xlMinuteCode             , "xlMinuteCode"              ' 23  Minute symbol (m).
   P  xlMonthCode              , "xlMonthCode"               ' 20  Month symbol (m).
   P  xlMonthLeadingZero       , "xlMonthLeadingZero"        ' 41  True if a leading zero is displayed in months (when months are displayed as numbers).
   P  xlMonthNameChars         , "xlMonthNameChars"          ' 30  Always returns three characters for backward compatibility. Abbreviated month names are read from Microsoft Windows and can be any length.
   P  xlNoncurrencyDigits      , "xlNoncurrencyDigits"       ' 29  Number of decimal digits to be used in noncurrency formats.
   P  xlNonEnglishFunctions    , "xlNonEnglishFunctions"     ' 34  True if you are not displaying functions in English.
   P  xlRightBrace             , "xlRightBrace"              ' 13  Character used instead of the right brace (}) in array literals.
   P  xlRightBracket           , "xlRightBracket"            ' 11  Character used instead of the right bracket (]) in R1C1-style references.
   P  xlRowSeparator           , "xlRowSeparator"            ' 15  Character used to separate rows in array literals.
   P  xlSecondCode             , "xlSecondCode"              ' 24  Second symbol (s).
   P  xlThousandsSeparator     , "xlThousandsSeparator"      '  4  Zero or thousands separator.
   P  xlTimeLeadingZero        , "xlTimeLeadingZero"         ' 45  True if a leading zero is displayed in times.
   P  xlTimeSeparator          , "xlTimeSeparator"           ' 18  Time separator (:).
   P  xlUpperCaseColumnLetter  , "xlUpperCaseColumnLetter"   '  7  Uppercase column letter.
   P  xlUpperCaseRowLetter     , "xlUpperCaseRowLetter"      '  6  Uppercase row letter (for R1C1-style references).
   P  xlWeekdayNameChars       , "xlWeekdayNameChars"        ' 31  Always returns three characters for backward compatibility. Abbreviated weekday names are read from Microsoft Windows and can be any length.
   P  xlYearCode               , "xlYearCode"                ' 19  Year symbol in number formats (y).

   columns(1).autofit
   columns(2).autofit

end sub ' }

sub P(xlVal as long, xlStr as string) ' {

    r = r + 1
    cells(r,1) = xlStr
    cells(r,2) = application.international(xlVal)

end sub ' }
