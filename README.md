# Ranges

> Build, convert and transform between range representations for Google Sheets.

[![Build Status](https://travis-ci.org/plemont/ranges.svg?branch=master)](https://travis-ci.org/plemont/ranges)
[![codecov](https://codecov.io/gh/plemont/ranges/branch/master/graph/badge.svg)](https://codecov.io/gh/plemont/ranges)

**Ranges**, or contiguous groups of cells, within Google Sheets can be represented in different ways, and furthermore can be either bounded or unbounded. Sometimes, the format you *need* isn't the format you *have*, or perhaps not the one easily calculated at that point in your application. This is where the **Ranges** library aims to help.

## Installation

**Ranges** is available from [Maven Central](https://search.maven.org/#artifactdetails%7Cio.github.plemont%7Cranges%7C1.0%7Cjar). Add to your `build.gradle`:

```
implementation 'io.github.plemont:ranges:1.0'
```

or with Maven:

```xml
<dependency>
    <groupId>io.github.plemont</groupId>
    <artifactId>ranges</artifactId>
    <version>1.0</version>
</dependency>
```

## Usage example

```java
// Build a A1 notation range from 0-indexed coordinates
String range = Ranges.forSheetName("Accounts")
    .withStartColumn(0)
    .withStartRow(0)
    .withEndColumn(9)
    .withEndRow(9)
    .toRange();  // Accounts!A1:J10.
    
String range = Ranges.forGridRange(gridRange)
    .withSheetName("Accounts")
    .toRange(); // Creates a A1 notation from a GridRange object.
    
String gridRange = Ranges
    .forRange("Accounts!C:D")
    .withSheetId(0)
    .toGridRange(); // Creates a GridRange from an A1 notation range.
    
String range = Ranges.forStartGridCoordinate(coord)
    .withWidth(10)
    .withHeight(10)
    .toRange(); // Creates an A1 notation range of specified extent, for a start GridCoordinate.
    
String range = Ranges.forRange("Test!A1:B2")
    .translate(5, 5)
    .toRange(); // Translates a range to Test!F6:G7
```

For further details on the transformations and conversions possible, see the API documentation.

## Release History

* 1.0 Initial release

## License

Distributed under the Apache 2.0 License. See [LICENSE](LICENSE) for more information.

## Contributing

1. Fork it (<https://github.com/plemont/ranges/fork>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request
