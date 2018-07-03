package io.github.plemont.ranges;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.common.collect.ImmutableList;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Provides conversion and manipulation functionality for ranges used in Google Sheets.
 *
 * <p>A range of cells can be represented in a number of ways when working with Google Sheets, be
 * it through the API, the UI or working programmatically with data structures in a client
 * application.
 *
 * <p>For example, the 5 by 5 grid at the origin of a sheet could be represented as:
 *
 * <ul>
 *   <li>A string - "Sheet1!A1:E5"</li>
 *   <li>A {@code GridRange}
 *     <pre>
 * {@code {
 *    sheetId: 0,
 *    startRowIndex: 0,
 *    endRowIndex: 5,
 *    startColumnIndex: 0,
 *    endColumnIndex: 5
 *  } }
 *     </pre>
 *   </li>
 * </ul>
 *
 * <h2>Converting between representations</h2>
 *
 * <p>Ranges can be built up from a number of start points using the {@code Ranges.for...}
 * methods, for example, the following will all output a range String, but from different start
 * points:
 *
 * <pre>
 * {@code
 *    String range = Ranges.forSheetName("Test").withStartCell("A1").withEndCell("C5").toRange();
 *    // range: Test!A1:C5.
 *
 *    String range = Ranges.forSheet(sheet).toRange();
 *    // Creates a range from a Google Sheet object.
 *
 *    String range = Ranges.forGridRange(gridRange).withSheetName("Test").toRange();
 *    // Creates a range from a GridRange objject.
 *
 *    String range = Ranges.forStartGridCoordinate(coord).withWidth(10).withHeight(10).toRange();
 *    // Creates a range starting at the specified cell, of width and height 10.
 *
 *    String range = Ranges.forRange("Test!A1:B2").translate(5, 5).toRange();
 *    // Creates a range from a range string and translates
 * }
 * </pre>
 */
public class Ranges {
  private static final int SHEET_NAME_MAX_LENGTH = 100;
  private static final int ASCII_A_OFFSET = 65;
  private static final int ALPHABET_LENGTH = 26;

  // Private constructor to avoid instantiation.
  private Ranges() {}

  /**
   * Inner context class, used for building and modifying the range prior to output in the desired
   * format.
   */
  static class RangeContext {
    private static final Pattern CELL_PATTERN = Pattern.compile("([A-Z]*)([0-9]*)");

    private String sheetName;
    private Integer sheetId;
    private Integer startColumn;
    private Integer startRow;
    private Integer endColumn;
    private Integer endRow;

    /**
     * Sets or overwrites the {@link com.google.api.services.sheets.v4.model.Sheet Sheet} name
     * for this context.
     *
     * <p>The supplied {@code sheetName} is tested for validity, See
     * {@link #isValidSheetName(String)} for details of what formats for {@code sheetName} are
     * accepted.
     *
     * @param sheetName The desired name for the {@code Sheet}.
     * @return the {@code RangeContext} object for chaining.
     */
    public RangeContext withSheetName(String sheetName) {
      checkArgument(isValidSheetName(sheetName),
              String.format("sheetName must be between 1 and %d characters.",
                  SHEET_NAME_MAX_LENGTH));
      this.sheetName = sheetName;
      return this;
    }

    /**
     * Sets of overwrites the {@code Sheet} ID for this context.
     *
     * @param sheetId The desired ID for the {@code Sheet}.
     * @return the {@code RangeContext} object for chaining.
     */
    public RangeContext withSheetId(int sheetId) {
      checkArgument(sheetId >= 0, "Sheet ID must be non-negative.");
      this.sheetId = sheetId;
      return this;
    }

    /**
     * Sets the width of the range, taken from the anchored start column.
     *
     * @param width The number of columns wide for the range. For example {@code 3} would include
     *     columns A,B and C, if the start column is set to A.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if the start column is not set.
     */
    public RangeContext withWidth(int width) {
      checkArgument(width > 0, "Width must be positive.");
      if (startColumn != null) {
        endColumn = startColumn + width - 1;
      } else {
        throw new IllegalStateException("Cannot set width where startColumn not set.");
      }
      return this;
    }

    /**
     * Sets the height of the range, taken from the anchored start row.
     *
     * @param height The number of rows for the range. For example {@code 3} would include rows 1,
     *     2 and 3, if the start row is set to row 1.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if the start row is not set.
     */
    public RangeContext withHeight(int height) {
      checkArgument(height > 0, "Height must be positive.");
      if (startRow != null) {
        endRow = startRow + height - 1;
      } else {
        throw new IllegalStateException("Cannot set height where range is not anchored startRow.");
      }
      return this;
    }

    /**
     * Sets the starting column for the range.
     *
     * <p>As well as setting the value for the {@code startColumn}, where both the
     * {@code startColumn} and {@code endColumn} are set, these are checked such that
     * {@code startColumn <= endColumn}.</p>
     * @param startColumn The zero-indexed start column.
     * @return the {@code RangeContext} object for chaining.
     */
    public RangeContext withStartColumn(int startColumn) {
      this.startColumn = startColumn;
      orderBounds();
      return this;
    }

    /**
     * Sets the starting row for the range.
     *
     * <p>As well as setting the value for the {@code startRow}, where both the
     * {@code startRow} and {@code endRow} are set, these are checked such that
     * {@code startRow <= endRow}.</p>
     * @param startRow The zero-indexed start row.
     * @return the {@code RangeContext} object for chaining.
     */
    public RangeContext withStartRow(int startRow) {
      this.startRow = startRow;
      orderBounds();
      return this;
    }

    /**
     * Sets the end column for the range.
     *
     * <p>As well as setting the value for the {@code endColumn}, where both the
     * {@code startColumn} and {@code endColumn} are set, these are checked such that
     * {@code startColumn <= endColumn}.</p>
     * @param endColumn The zero-indexed end column, inclusive. e.g. {@code 3} sets the fourth
     *     column as the last column included in the range.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if the start column is not set.
     */
    public RangeContext withEndColumn(int endColumn) {
      if (startColumn != null) {
        this.endColumn = endColumn;
      } else {
        throw new IllegalStateException("Cannot set endColumn where startColumn not set.");
      }
      orderBounds();
      return this;
    }

    /**
     * Sets the end row for the range.
     *
     * <p>As well as setting the value for the {@code endRow}, where both the
     * {@code startRow} and {@code endRow} are set, these are checked such that
     * {@code startRow <= endRow}.</p>
     * @param endRow The zero-indexed end row, inclusive. e.g. {@code 3} sets the fourth
     *     row as the last row included in the range.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if the start row is not set.
     */
    public RangeContext withEndRow(int endRow) {
      if (startRow != null) {
        this.endRow = endRow;
      } else {
        throw new IllegalStateException("Cannot set endRow where startRow not set.");
      }
      orderBounds();
      return this;
    }

    /**
     * Sets the start cell for a range.
     *
     * @param a1Cell The start cell for the range, in A1 notation, e.g. A1 represents the cell
     *     (0, 0). Where {@code endColumn} and/or {@code endRow} has been set, either individually
     *     or through {@code withEndCell}, ordering is checked to ensure that
     *     {@code startColumn <= endColumn} and {@code startRow <= endRow}.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalArgumentException if the specified cell format is invalid.
     */
    public RangeContext withStartCell(String a1Cell) {
      checkNotNull(a1Cell, "start cell cannot be null");
      checkArgument(a1Cell.length() > 0, "start cell cannot be empty string");
      Matcher matcher = CELL_PATTERN.matcher(a1Cell);
      if (matcher.matches()) {
        if (matcher.group(1).length() > 0) {
          this.startColumn = alphaColumnToColumnIndex(matcher.group(1)) - 1;
        }
        if (matcher.group(2).length() > 0) {
          int row = Integer.parseInt(matcher.group(2)) - 1;
          if (row < 0) {
            throw new IllegalArgumentException("Invalid row specified");
          }
          this.startRow = row;
        }
      } else {
        throw new IllegalArgumentException("Illegal cell format.");
      }
      orderBounds();
      return this;
    }

    /**
     * Sets the end cell for a range.
     *
     * @param a1Cell The end cell for the range, in A1 notation, e.g. A1 represents the cell
     *     (0, 0). Column and row ordering is then checked to ensure that
     *     {@code startColumn <= endColumn} and {@code startRow <= endRow}.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalArgumentException if the specified cell format is invalid.
     */
    public RangeContext withEndCell(String a1Cell) {
      checkNotNull(a1Cell, "end cell cannot be null");
      checkArgument(a1Cell.length() > 0, "end cell cannot be empty string");
      Matcher matcher = CELL_PATTERN.matcher(a1Cell);
      if (matcher.matches()) {
        if (matcher.group(1).length() > 0) {
          if (this.startColumn !=  null) {
            this.endColumn = alphaColumnToColumnIndex(matcher.group(1)) - 1;
          } else {
            throw new IllegalStateException("Cannot set endColumn when startColumn is unset.");
          }
        }
        if (matcher.group(2).length() > 0) {
          int row = Integer.parseInt(matcher.group(2)) - 1;
          if (row < 0) {
            throw new IllegalArgumentException("Invalid row specified");
          } else if (this.startRow == null) {
            throw new IllegalStateException("Cannot set endRow bound when startRow is unset.");
          }
          this.endRow = row;
        }
      } else {
        throw new IllegalArgumentException("Illegal cell format.");
      }
      orderBounds();
      return this;
    }

    /**
     * Clears the start column. Used for creating unbounded ranges from bounded ranges.
     *
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if attempting to unset the start column when end column is set.
     */
    public RangeContext clearStartColumn() {
      if (endColumn != null) {
        throw new IllegalStateException("Cannot clear startColumn where endColumn still set.");
      }
      this.startColumn = null;
      return this;
    }

    /**
     * Clears the start row. Used for creating unbounded ranges from bounded ranges.
     *
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if attempting to unset the start row when end row is set.
     */
    public RangeContext clearStartRow() {
      if (endRow != null) {
        throw new IllegalStateException("Cannot clear startRow where endRow still set.");
      }
      this.startRow = null;
      return this;
    }

    /**
     * Clears the end column. Used for creating unbounded ranges from bounded ranges.
     *
     * @return the {@code RangeContext} object for chaining.
     */
    public RangeContext clearEndColumn() {
      this.endColumn = null;
      return this;
    }

    /**
     * Clears the end row. Used for creating unbounded ranges from bounded ranges.
     *
     * @return the {@code RangeContext} object for chaining.
     */
    public RangeContext clearEndRow() {
      this.endRow = null;
      return this;
    }

    /**
     * Expands the height of the range by the specified number of rows.
     *
     * @param numExtraRows The number of extra rows to add to the range.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if attempting to expand a range with no start or end rows.
     */
    public RangeContext expandRows(int numExtraRows) {
      checkArgument(numExtraRows > 0, "numExtraRows must be greater than zero.");
      if (startRow == null || endRow == null) {
        throw new IllegalStateException("Cannot expand rows where bounds are not set.");
      }
      endRow += numExtraRows;
      return this;
    }

    /**
     * Expands the width of the range by the specified number of columns.
     *
     * @param numExtraColumns The number of extra rows to add to the range.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if attempting to expand a range with no start or end columns.
     */
    public RangeContext expandColumns(int numExtraColumns) {
      checkArgument(numExtraColumns > 0, "numExtraColumns must be greater than zero.");
      if (startColumn == null || endColumn == null) {
        throw new IllegalStateException("Cannot expand columns where bounds are not set.");
      }
      endColumn += numExtraColumns;
      return this;
    }

    /**
     * Translates a range by the specified (x, y) adjustment.
     *
     * <p>Translating a range requires that the start column or row is set for any non-zero
     * x or y-axis adjusts respectively.</p>
     *
     * @param deltaX The number of columns to translate by.
     * @param deltaY The number of rows to translate by.
     * @return the {@code RangeContext} object for chaining.
     * @throws IllegalStateException if attempting to translate along an axis where no start is set.
     * @throws IllegalArgumentException if the delta would translate to a negative column or row.
     */
    public RangeContext translate(int deltaX, int deltaY) {
      if (deltaX != 0) {
        if (startColumn == null) {
          throw new IllegalStateException("Cannot translate range where startColumn is not set.");
        } else if (startColumn + deltaX < 0) {
          throw new IllegalArgumentException("Cannot translate to before column 0.");
        }
        startColumn += deltaX;
        if (endColumn != null) {
          endColumn += deltaX;
        }
      }
      if (deltaY != 0) {
        if (startRow == null) {
          throw new IllegalStateException("Cannot translate range where startRow is not set.");
        } else if (startRow + deltaY < 0) {
          throw new IllegalArgumentException("Cannot translate to before row 0.");
        }
        startRow += deltaY;
        if (endRow != null) {
          endRow += deltaY;
        }
      }
      return this;
    }

    /**
     * Forms a range String from the current context.
     *
     * <p>Range strings identify an area of a given {@code Sheet}, as specifically as a single
     * cell, or to as generally as the whole {@code Sheet}, for example:
     * </p>
     *
     * <ul>
     *   <li><strong>Sheet1</strong> - Basic range, specifying the entire {@code Sheet} named
     *       Sheet1.</li>
     *   <li><strong>Sheet1!A1</strong> - The cell (0, 0) for Sheet1.</li>
     *   <li><strong>Sheet1!A1:C4</strong> - The rectangular area from (0, 0) to (2, 3)
     *       inclusive.</li>
     *   <li><strong>Sheet1!A:C</strong> - The first three columns in Sheet1 inclusive.</li>
     *   <li><strong>Sheet1!2:6</strong> - The second to sixth rows in Sheet1 inclusive.</li>
     *   <li><strong>Sheet1:A4:E</strong> - The first five columns, starting at row 4.</li>
     *   <li><strong>'Today''s data'!A1</strong> - A range for {@code Sheet} with
     *       name "Today's data", using single-quote escaping.</li>
     * </ul>
     *
     * @return the range string.
     * @throws IllegalStateException if the current context cannot be converted to a range String.
     */
    public String toRange()  {
      if (sheetName == null) {
        throw new IllegalStateException("Sheet name is not set: cannot create a range string.");
      }
      String range = escapeSheetName(sheetName);

      // Full-grid specification
      if (startColumn != null && endColumn != null && startRow != null && endRow != null) {
        String startCell = columnIndexToAlphaColumn(startColumn) + (startRow + 1);
        String endCell = columnIndexToAlphaColumn(endColumn) + (endRow + 1);
        range += "!" + startCell;
        if (!startCell.equals(endCell)) {
          range += ":" + endCell;
        }
      } else if (startColumn != null && endColumn != null && endRow == null) {
        String startRowString = startRow == null ? "" : String.valueOf(startRow + 1);
        range += String.join("", "!", columnIndexToAlphaColumn(startColumn), startRowString, ":",
            columnIndexToAlphaColumn(endColumn));
      } else if (endColumn == null && startRow != null && endRow != null) {
        String startColumnString = startColumn == null ? "" : columnIndexToAlphaColumn(startColumn);
        range += String.join("", "!", startColumnString, String.valueOf(startRow + 1),
            ":", String.valueOf(endRow + 1));
      } else if (startColumn != null || endColumn != null || startRow != null || endRow != null) {
        throw new IllegalStateException("Illegal combination of coordinates set.");
      }
      return range;
    }

    /**
     * Creates a {@link com.google.api.services.sheets.v4.model.GridRange GridRange} from the
     * current context.
     *
     * @return the created GridRange.
     */
    public GridRange toGridRange() {
      GridRange gridRange = new GridRange();
      gridRange.setSheetId(sheetId);
      gridRange.setStartRowIndex(startRow);
      gridRange.setEndRowIndex(endRow != null ? endRow + 1 : null);
      gridRange.setStartColumnIndex(startColumn);
      gridRange.setEndColumnIndex(endColumn != null ? endColumn + 1 : null);
      return gridRange;
    }

    /**
     * Creates a {@link com.google.api.services.sheets.v4.model.GridCoordinate GridCoordinate} from
     * the start cell of the current context.
     *
     * @return the created GridCoordinate.
     */
    public GridCoordinate toStartGridCoordinate() {
      GridCoordinate gridCoordinate = new GridCoordinate();
      gridCoordinate.setSheetId(sheetId);
      gridCoordinate.setColumnIndex(startColumn);
      gridCoordinate.setRowIndex(startRow);
      return gridCoordinate;
    }

    /**
     * Creates a {@link com.google.api.services.sheets.v4.model.GridCoordinate GridCoordinate} from
     * the end cell of the current context.
     *
     * @return the created GridCoordinate.
     */
    public GridCoordinate toEndGridCoordinate() {
      GridCoordinate gridCoordinate = new GridCoordinate();
      gridCoordinate.setSheetId(sheetId);
      gridCoordinate.setColumnIndex(endColumn);
      gridCoordinate.setRowIndex(endRow);
      return gridCoordinate;
    }

    /**
     * Retrieve the name of the {@code Sheet} in the current context.
     *
     * @return the sheet name.
     */
    public String getSheetName() {
      return sheetName;
    }

    /**
     * Ensures that {@code startRow <= endRow} and {@code startColumn <= endColumn} for the current
     * context, by swapping values where necessary when both start and end are defined.
     */
    private void orderBounds() {
      if (startColumn != null && endColumn != null) {
        int a = startColumn;
        int b = endColumn;
        startColumn = Math.min(a, b);
        endColumn = Math.max(a, b);
      }
      if (startRow != null && endRow != null) {
        int a = startRow;
        int b = endRow;
        startRow = Math.min(a, b);
        endRow = Math.max(a, b);
      }
    }

    /**
     * Encloses a {@code Sheet} name with single-quotes and escapes any single-quotes where
     * necessary.
     *
     * <p>{@code Sheet} names consisting only of alphanumeric characters require no escaping,
     * however, any punctuation will result in the entire name being enclosed in quotes. Some
     * examples:</p>
     *
     * <ul>
     *   <ul><strong>Sheet1</strong> - is escaped as {@code Sheet1}</ul>
     *   <ul><strong>My Sheet</strong> - is escaped as {@code 'My Sheet'}</ul>
     *   <ul><strong>Today's data</strong> - is escaped as {@code 'Today''s data'}</ul>
     * </ul>
     *
     * @param sheetName The string to escape.
     * @return the escaped sheet name.
     */
    private String escapeSheetName(String sheetName) {
      if (sheetName.matches("[A-Za-z0-9]+")) {
        return sheetName;
      }
      return "'" + sheetName.replaceAll(Pattern.quote("'"), "''") + "'";
    }

    /**
     * Determines whether a String represents a valid {@code Sheet} name.
     *
     * @param sheetName The String to check.
     * @return whether or not the name is valid.
     */
    private boolean isValidSheetName(String sheetName) {
      return sheetName != null
          && sheetName.length() > 0 && sheetName.length() < SHEET_NAME_MAX_LENGTH;
    }

    /**
     * Converts a numeric column index into a String representation as per A1 cell notation. For
     * example: {@code 0 -> A}.
     *
     * @param columnIndex The 0-indexed column number.
     * @return The A1-format column specifier.
     */
    private String columnIndexToAlphaColumn(int columnIndex) {
      checkArgument(columnIndex >= 0, "Column index should be greater or equal to zero.");
      StringBuilder sb = new StringBuilder();

      int remainder = columnIndex % ALPHABET_LENGTH;
      sb.append((char) (remainder + ASCII_A_OFFSET));

      while (columnIndex - remainder > 0) {
        columnIndex = (columnIndex - remainder) / ALPHABET_LENGTH - 1;
        remainder = columnIndex % ALPHABET_LENGTH;
        sb.insert(0, (char) (remainder + ASCII_A_OFFSET));
      }
      return sb.toString();
    }
  }


  /**
   * Creates a new {@code RangeContext} for manipulating ranges based on a sheet name as a start
   * point.
   *
   * @param sheetName The sheet name.
   * @return the {@code RangeContext} object for chaining.
   */
  public static RangeContext forSheetName(String sheetName) {
    RangeContext rangeContext = new RangeContext();
    return rangeContext.withSheetName(sheetName);
  }

  /**
   * Creates a new {@code RangeContext} for manipulating ranges based on a {code Sheet} object.
   *
   * @param sheet The {@code Sheet} object.
   * @return the {@code RangeContext} object for chaining.
   */
  public static RangeContext forSheet(Sheet sheet) {
    checkNotNull(sheet, "sheet cannot be null.");
    RangeContext rangeContext = new RangeContext();
    SheetProperties props = sheet.getProperties();
    return rangeContext.withSheetName(props.getTitle())
            .withSheetId(props.getSheetId());
  }

  /**
   * Creates a new {@code RangeContext} for manipulating ranges based on a {code GridRange} object.
   *
   * @param gridRange The {@code GridRange} object.
   * @return the {@code RangeContext} object for chaining.
   */
  public static RangeContext forGridRange(GridRange gridRange) {
    checkNotNull(gridRange, "gridrange cannot be null.");
    RangeContext rangeContext = new RangeContext();
    rangeContext.withSheetId(gridRange.getSheetId());
    rangeContext.withStartColumn(gridRange.getStartColumnIndex());
    rangeContext.withStartRow(gridRange.getStartRowIndex());
    rangeContext.withEndColumn(gridRange.getEndColumnIndex() - 1);
    rangeContext.withEndRow(gridRange.getEndRowIndex() - 1);
    return rangeContext;
  }

  /**
   * Creates a new {@code RangeContext} for manipulating ranges based on a {code GridCoordinate}
   * object.
   *
   * @param gridCoordinate The {@code GridCoordinate} object to use as the start cell.
   * @return the {@code RangeContext} object for chaining.
   */
  public static RangeContext forStartGridCoordinate(GridCoordinate gridCoordinate) {
    checkNotNull(gridCoordinate, "gridCoordinate cannot be null.");
    RangeContext rangeContext = new RangeContext();
    rangeContext.withSheetId(gridCoordinate.getSheetId());
    rangeContext.withStartColumn(gridCoordinate.getColumnIndex());
    rangeContext.withStartRow(gridCoordinate.getRowIndex());
    return rangeContext;
  }

  /**
   * Creates a new {@code RangeContext} for manipulating ranges based on a range String.
   *
   * <p>See {@link RangeContext#toRange()} for details of valid range String formats.</p>
   *
   * @param range The range string.
   * @return the {@code RangeContext} object for chaining.
   * @throws IllegalArgumentException if the range is invalid.
   */
  public static RangeContext forRange(String range) {
    // Sheet names appear to be able to consist of pretty much any character. When including a
    // special character such as a space, comma, dash etc, the name must be quoted.
    String simpleSheetNameRx = String.format("[A-Za-z0-9]{1,%d}", SHEET_NAME_MAX_LENGTH);
    String spacedSheetNameRx = String.format("'(?:''|[\\x20-\\x26\\x28-\\x7E]){1,%d}'",
        SHEET_NAME_MAX_LENGTH);

    // This is a basic regex for capturing the optional 1 or 2 A1-notation coordinates in the range
    // specification. This regex is over-generous, allowing for invalid cases in the name of
    // simplicity, which are then tested for further below and filtered out.
    String gridRx = "(?:|!([A-Z]*)([0-9]*)(?:|(:)([A-Z]*)([0-9]*)))";

    String pattern = String.format("(%s|%s)%s", simpleSheetNameRx, spacedSheetNameRx, gridRx);
    Pattern rangePattern = Pattern.compile(pattern);

    Matcher matcher = rangePattern.matcher(range);
    if (matcher.matches()) {
      List<Integer> coords = ImmutableList.of(getColIntFromExtractedString(matcher.group(2)),
          getRowIntFromExtractedString(matcher.group(3)),
          getColIntFromExtractedString(matcher.group(5)),
          getRowIntFromExtractedString(matcher.group(6)));

      boolean isColon = matcher.group(4) != null && matcher.group(4).equals(":");
      coords = checkCoordEdgeCases(coords, isColon);
      coords = checkCoordOrdering(coords);

      String sheetName = unescapeSheetName(matcher.group(1));
      return sheetNameAndCoordsToRangeContext(sheetName, coords);
    } else {
      throw new IllegalArgumentException("Not a valid range.");
    }
  }

  /**
   * Convenience method to create a {@code RangeContext} for a given sheet name and set of four
   * coordinates.
   *
   * @param sheetName The desired {@code Sheet} name.
   * @param coords A list of coordinates, in the form [startColumn, startRow, endColumn, endRow].
   *     Note that these coordinates are 1-indexed, as 0 is the case where the parameter is not
   *     specified, and should be left as null in the {@code RangeContext}.
   * @return the {@code RangeContext} object for chaining.
   */
  private static RangeContext sheetNameAndCoordsToRangeContext(String sheetName,
      List<Integer> coords) {
    RangeContext rangeContext = new RangeContext();
    rangeContext.withSheetName(sheetName);
    if (coords.get(0) > 0) {
      rangeContext.withStartColumn(coords.get(0) - 1);
    }
    if (coords.get(1) > 0) {
      rangeContext.withStartRow(coords.get(1) - 1);
    }
    if (coords.get(2) > 0) {
      rangeContext.withEndColumn(coords.get(2) - 1);
    }
    if (coords.get(3) > 0) {
      rangeContext.withEndRow(coords.get(3) - 1);
    }
    return rangeContext;
  }

  /**
   * Checks validity edge cases for coordinates supplied as part of a range.
   *
   * @param coords A list of coordinates, in the form [startColumn, startRow, endColumn, endRow].
   *     Note that these coordinates are 1-indexed, as 0 is the case where the parameter is not
   *     specified, and should be left as null in the {@code RangeContext}.
   * @param isColon Whether a colon was present in the range, to separate start and end cell.
   * @return a new list if it is valid.
   * @throws IllegalArgumentException if the range is invalid.
   */
  private static List<Integer> checkCoordEdgeCases(List<Integer> coords, boolean isColon) {
    // Edge-case 1: If there is a dividing ":" but no second coordinate specified
    // e.g. Sheet1!A1: throw an error:
    if (isColon && coords.get(2) == 0 && coords.get(3) == 0) {
      throw new IllegalArgumentException("Colon in range but no second coordinate specified.");
    }

    // Edge-case 2: Check for 3 occurrences of unspecified coordinate part.
    if (Collections.frequency(coords, 0) == 3) {
      throw new IllegalArgumentException("Single-dimension range coords not valid in isolation.");
    }

    // Edge-case 3: A pair of range coords, separated by ":" where one is just row, the other just
    // column, is invalid.
    if ((coords.get(0) > 0 && coords.get(1) == 0 && coords.get(2) == 0 && coords.get(3) > 0)
        || (coords.get(0) == 0 && coords.get(1) > 0 && coords.get(2) > 0 && coords.get(3) == 0)) {
      throw new IllegalArgumentException("Ranges cannot consist of <row>:<col> or <col>:<row>.");
    }

    // Edge-case 4: If one coordinate is set with both row and column, it means it's a bounded 1x1
    // range. Set the end of the range accordingly.
    if (coords.get(0) > 0 && coords.get(1) > 0 && coords.get(2) == 0 && coords.get(3) == 0) {
      return ImmutableList.of(coords.get(0), coords.get(1), coords.get(0), coords.get(1));
    }
    return coords;
  }

  /**
   * Re-orders coordinates to ensure that for both columns and rows, start is less than or equal to
   * end.
   *
   * <p>For example coordinates [2, 2, 1, 1] should be re-ordered to [1, 1, 2, 2] (representing
   * that B2:A1 in A1 notation should be reordered to A1:B2).
   *
   * @param coords A list of coordinates, in the form [startColumn, startRow, endColumn, endRow].
   *     Note that these coordinates are 1-indexed, as 0 is the case where the parameter is not
   *     specified, and should be left as null in the {@code RangeContext}.
   * @return The re-ordered list.
   */
  private static List<Integer> checkCoordOrdering(List<Integer> coords) {
    // Test for ordering: startColumn should be <= endColumn, startRow should be <= endRow
    int startColumn = coords.get(0);
    int endColumn = coords.get(2);
    if (endColumn != 0 && (startColumn == 0 || endColumn < startColumn)) {
      int temp = startColumn;
      startColumn = endColumn;
      endColumn = temp;
    }
    int startRow = coords.get(1);
    int endRow = coords.get(3);
    if (endRow != 0 && (startRow == 0 || endRow < startRow)) {
      int temp = startRow;
      startRow = endRow;
      endRow = temp;
    }
    return ImmutableList.of(startColumn, startRow, endColumn, endRow);
  }

  /**
   * Converts the column representation from a String to a number.
   *
   * @param matchedString The numeric string.
   * @return the extracted integer.
   */
  private static int getColIntFromExtractedString(String matchedString) {
    if (matchedString == null || matchedString.length() == 0) {
      return 0;
    }
    return alphaColumnToColumnIndex(matchedString);
  }

  /**
   * Converts the row representation from an A1 notation to a 1-index integer. For example, A1 -> 1.
   *
   * @param matchedString The numeric string.
   * @return the extracted integer.
   * @throws IllegalArgumentException if the number is not > 0.
   */
  private static int getRowIntFromExtractedString(String matchedString) {
    if (matchedString == null || matchedString.length() == 0) {
      return 0;
    }
    int row = Integer.parseInt(matchedString);
    if (row <= 0) {
      throw new IllegalArgumentException("Row must be a positive integer >= 1");
    }
    return row;
  }

  /**
   * Converts the alpha representation to numeric for a column index. For example A -> 1.
   *
   * @param a1Column The string representation, e.g. "A" or "AA".
   * @return the resulting integer, 1-indexed.
   */
  private static int alphaColumnToColumnIndex(String a1Column) {
    checkNotNull(a1Column, "a1Column cannot be null.");
    checkArgument(a1Column.length() > 0, "a1Column must be one character or longer.");
    int columnIndex = 0;
    for (int i = 0; i < a1Column.length(); i++) {
      int c = a1Column.charAt(i) - ASCII_A_OFFSET + 1;
      columnIndex += c * Math.pow(ALPHABET_LENGTH, (double)(a1Column.length() - 1 - i));
    }
    return columnIndex;
  }

  /**
   * Removes the escaping from the sheet part of a range string.
   *
   * <p>For escaping rules see {@link RangeContext#escapeSheetName(String)}.</p>
   *
   * @param escapedSheetName The escaped sheet name.
   * @return the unescaped sheet name.
   */
  private static String unescapeSheetName(String escapedSheetName) {
    if (escapedSheetName.startsWith("'") && escapedSheetName.endsWith("'")) {
      return escapedSheetName.substring(1, escapedSheetName.length() - 1)
          .replaceAll(Pattern.quote("''"), "'");
    }
    return escapedSheetName;
  }
}
