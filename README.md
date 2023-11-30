# Number Formatters Library for ASP

VBScript includes the following functions: FormatNumber(), FormatCurrency(), FormatPercentage(). PHP's NumberFormatter class additionally offers the ability to format scientific notation, ordinals, and spellout. This library adds equivalent functions for use in ASP: FormatScientific(), FormatOrdinal(), FormatSpellout(), and FormatSpelloutOrdinal().

## Project Status

No further development is currently planned, but a potential next step would be to consolidate these new functions along with the built-in functions into a NumberFormatter class similar to the one found in PHP.

## Installation

Place NumberFormatters.lib.asp in any location on your web server, or on another machine on the same network. For additional security, you may wish to place it in a location that isn't directly accessible by users.

The file NumberFormatters.test.asp is not required in order to use the library and does not need to be placed on the web server unless you want to run unit tests.

## Usage

```vbscript
dim cardinal
cardinal = 1234567890

' Format a number using scientific notation.
dim scientific
scientific = FormatScientific(cardinal)

' Format a number as an ordinal.
dim ordinal
ordinal = FormatOrdinal(cardinal)

' Spellout a cardinal number.
dim spellout
spellout = FormatSpellout(cardinal)

' Spellout an ordinal number.
dim spellout-ordinal
spellout-ordinal = FormatSpelloutOrdinal(cardinal)
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

See NumberFormatters.test.asp for unit tests.

## Authors

Version 1.0 written April/September 2008 by Scott Vander Molen

Version 2.0 written November 2023 by Scott Vander Molen

## License
This work is licensed under a [Creative Commons Attribution 4.0 International License](https://creativecommons.org/licenses/by/4.0/).