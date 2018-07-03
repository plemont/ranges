package io.github.plemont.ranges;

import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.GridRange;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import java.lang.reflect.Constructor;
import org.junit.Before;
import org.junit.Test;

import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.Assert.*;

public class RangesTest {
  private static final int SHEET_NAME_MAX_LENGTH = 100;

  @Before
  public void setUp() {

  }

  @Test
  public void constructorTest() throws Exception {
    Constructor[] constructors = Ranges.class.getDeclaredConstructors();
    assertEquals(1, constructors.length);
    Constructor constructor = constructors[0];
    assertFalse(constructor.isAccessible());
    constructor.setAccessible(true);
    assertEquals(Ranges.class, constructor.newInstance().getClass());
  }

  @Test
  public void forSheetName_validName() {
    String expected = "TestSheet1";
    assertEquals(expected, Ranges.forSheetName("TestSheet1").getSheetName());
  }

  @Test
  public void forSheetName_emptyName() {
    try {
      Ranges.forSheetName("").getSheetName();
      fail();
    } catch (RuntimeException e) {
      // Expected exception for empty sheet name.
    }
  }

  @Test
  public void forSheetName_nameTooLong() {
    String sheetName = IntStream.range(0, SHEET_NAME_MAX_LENGTH + 1)
            .mapToObj(i -> "a").collect(Collectors.joining(""));
    try {
      Ranges.forSheetName(sheetName).getSheetName();
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for too long a sheet name.
    }
  }

  @Test
  public void forSheetName_widthWithoutStartColumn() {
    try {
      Ranges.forSheetName("Test").withWidth(10).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Can't set the width when the start column has not been set.
    }
  }

  @Test
  public void forSheetName_widthLessThanOne() {
    try {
      Ranges.forSheetName("Test!A1").withWidth(0).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Can't set the width < 1
    }
  }

  @Test
  public void forSheetName_heightLessThanOne() {
    try {
      Ranges.forSheetName("Test!A1").withHeight(0).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Can't set the height < 1
    }
  }

  @Test
  public void forSheetName_heightWithoutStartRow() {
    try {
      Ranges.forSheetName("Test").withHeight(10).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Can't set the height when the start row has not been set.
    }
  }

  @Test
  public void forSheetName_invalidSheetId() {
    try {
      Ranges.forSheetName("Test").withSheetId(-1).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Sheet ID must be >= 0.
    }
  }

  @Test
  public void forSheet_nullValue() {
    try {
      Ranges.forSheet(null);
      fail();
    } catch (NullPointerException e) {
      // Expected exception for empty sheet object.
    }
  }

  @Test
  public void forRange_nullValue() {
    try {
      Ranges.forRange(null);
      fail();
    } catch (NullPointerException e) {
      // Expected exception for empty range string.
    }
  }

  @Test
  public void forRange_invalidName() {
    try {
      Ranges.forRange("Test!A:B:C:D");
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for an invalid range specification.
    }
  }

  @Test
  public void forRange_invalidRowColumnOnly() {
    try {
      Ranges.forRange("Test!A:5");
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for range where each of the coordinate pair is row and column
      // respectively only, e.g. Test!A:5 or Test!6:D
    }
  }

  @Test
  public void forRange_plainString() {
    assertEquals("sheet1", Ranges.forRange("sheet1").toRange());
  }

  @Test
  public void forRange_containsEscapedSingleQuote() {
    assertEquals("'Brian''s Sheet'", Ranges.forRange("'Brian''s Sheet'").toRange());
  }

  @Test
  public void forRange_fullGridSpecification() {
    assertEquals("'Brian''s Sheet'!A1:ZC1000", Ranges.forRange("'Brian''s Sheet'!A1:ZC1000").toRange());
  }

  @Test
  public void forRange_forRange_fullGridSpecificationZeroRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'!C0:D10");
      fail();
    } catch (IllegalArgumentException e) {
      // Row specifier must be > 0.
    }
  }

  @Test
  public void forRange_singleCell() {
    assertEquals("'Brian''s Sheet'!CD500", Ranges.forRange("'Brian''s Sheet'!CD500").toRange());
  }

  @Test
  public void forRange_singleCellGridRange() {
    GridRange range = Ranges.forRange("'Brian''s Sheet'!CD500").withSheetId(0).toGridRange();
    assertEquals(81, range.getStartColumnIndex().intValue());
    assertEquals(499, range.getStartRowIndex().intValue());
    assertEquals(82, range.getEndColumnIndex().intValue());
    assertEquals(500, range.getEndRowIndex().intValue());
  }

  @Test
  public void forRange_singleCellWithWidthAndHeightGridRange() {
    GridRange range = Ranges.forRange("'Brian''s Sheet'!CD500")
            .withSheetId(0)
            .withWidth(10)
            .withHeight(10)
            .toGridRange();
    assertEquals(81, range.getStartColumnIndex().intValue());
    assertEquals(499, range.getStartRowIndex().intValue());
    assertEquals(91, range.getEndColumnIndex().intValue());
    assertEquals(509, range.getEndRowIndex().intValue());
  }

  @Test
  public void forRange_invalidSingleCellZeroRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'!C0");
      fail();
    } catch (IllegalArgumentException e) {
      // Row specifier must be > 0.
    }
  }

  @Test
  public void forRange_columns() {
    assertEquals("'Brian''s Sheet'!C:D", Ranges.forRange("'Brian''s Sheet'!C:D").toRange());
  }

  @Test
  public void forRange_columnsReverseOrder() {
    assertEquals("'Brian''s Sheet'!C:D", Ranges.forRange("'Brian''s Sheet'!D:C").toRange());
  }

  @Test
  public void forRange_rows() {
    assertEquals("'Brian''s Sheet'!4:20", Ranges.forRange("'Brian''s Sheet'!4:20").toRange());
  }

  @Test
  public void forRange_rowsWithColumnAnchor() {
    assertEquals("'Brian''s Sheet'!B4:20", Ranges.forRange("'Brian''s Sheet'!B4:20").toRange());
  }

  @Test
  public void forRange_columnsWithRowAnchor() {
    assertEquals("'Brian''s Sheet'!B4:D", Ranges.forRange("'Brian''s Sheet'!B4:D").toRange());
  }

  @Test
  public void forRange_columnsWithRowAnchorReverseRow() {
    assertEquals("'Brian''s Sheet'!B4:D", Ranges.forRange("'Brian''s Sheet'!B:D4").toRange());
  }

  @Test
  public void forRange_startCell() {
    assertEquals("'Brian''s Sheet'!A10:D",
            Ranges.forRange("'Brian''s Sheet'!B:D4").withStartCell("A10").toRange());
  }

  @Test
  public void forRange_startCellEndCell() {
    assertEquals("'Brian''s Sheet'!A10:D",
            Ranges.forRange("'Brian''s Sheet'").withStartCell("A10").withEndCell("D").toRange());
  }

  @Test
  public void forRange_startCellZeroRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withStartCell("AC0").toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Row must be >= 1
    }
  }

  @Test
  public void forRange_startCellEndCellColumnOnly() {
    assertEquals("Test!A:B",
        Ranges.forSheetName("Test").withStartCell("A").withEndCell("B").toRange());
  }

  @Test
  public void forRange_startCellEndCellRowOnly() {
    assertEquals("Test!1:5",
        Ranges.forSheetName("Test").withStartCell("1").withEndCell("5").toRange());
  }

  @Test
  public void forRange_startCellEmptyString() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withStartCell("").toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Start cell cannot be empty string.
    }
  }

  @Test
  public void forRange_endCellEmptyString() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withStartCell("A1").withEndCell("").toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // End cell cannot be empty string.
    }
  }

  @Test
  public void forRange_startCellIllegalFormat() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withStartCell("AB0CEF").toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Format must be valid (col)(row).
    }
  }

  @Test
  public void forRange_endCellIllegalFormat() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withStartCell("A1").withEndCell("AE1D3").toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Format must be valid (col)(row).
    }
  }

  @Test
  public void forRange_startCellendCell() {
    assertEquals("'Today''s report'!A1:D10",
        Ranges.forRange("'Today''s report'")
            .withStartCell("D1").withEndCell("A10").toRange());
  }

  @Test
  public void forRange_endCellWithoutStartColumn() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withEndCell("C5").toRange();
      fail();
    } catch (IllegalStateException e) {
      // Start column must already be set if end column is to be set.
    }
  }

  @Test
  public void forRange_encCellZeroRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withEndCell("0").toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Row cannot be zero
    }
  }

  @Test
  public void forRange_endCellWithoutStartRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'").withEndCell("5").toRange();
      fail();
    } catch (IllegalStateException e) {
      // start cell must already be set if end cell is to be set.
    }
  }

  @Test
  public void forRange_columnsWithRowAnchorGridRange() {
    GridRange range = Ranges.forRange("'Brian''s Sheet'!B:D4").toGridRange();
    assertEquals(1, range.getStartColumnIndex().intValue());
    assertEquals(3, range.getStartRowIndex().intValue());
    assertEquals(4, range.getEndColumnIndex().intValue());
    assertNull(range.getEndRowIndex());
  }

  @Test
  public void forRange_invalidSingleCell() {
    try {
      Ranges.forRange("'Brian''s Sheet'!C");
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for range with only one part (no colon) but lacking both column and row.
    }
  }

  @Test
  public void forRange_incompleteRangePart() {
    try {
      Ranges.forRange("'Brian''s Sheet'!CD500:");
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for range with colon but no second part to range.
    }
  }

  @Test
  public void forRange_columnAndRowOnlyPair() {
    try {
      Ranges.forRange("'Brian''s Sheet'!A:10");
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for range of form ...!<row>:<column> or ...!<column>:<row>
    }
  }

  @Test
  public void forRange_setOnlyStartColumnRange() {
    try {
      Ranges.forRange("'Brian''s Sheet'!A:B").clearEndColumn().toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expect an exception when just the start column is set only.
    }
  }

  @Test
  public void forRange_setOnlyStartRowRange() {
    try {
      Ranges.forRange("'Brian''s Sheet'!2:20").clearEndRow().toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expect an exception when just the start row is set only.
    }
  }

  @Test
  public void forRange_clearColumnConstraintsRange() {
    assertEquals("Test!4:40",
        Ranges.forRange("Test!A4:C40").clearEndColumn().clearStartColumn().toRange());
  }

  @Test
  public void forRange_clearRowConstraintsRange() {
    assertEquals("Test!A:C",
        Ranges.forRange("Test!A1:C10").clearEndRow().clearStartRow().toRange());
  }

  @Test
  public void forRange_clearColumnConstraintsRangeWrongOrder() {
   try {
     Ranges.forRange("Test!A4:C40").clearStartColumn().clearEndColumn().toRange();
     fail();
   } catch (IllegalStateException e) {
     // expected exception from not clearing coordinates out of order.
   }
  }

  @Test
  public void forRange_clearRowonstraintsRangeWrongOrder() {
    try {
      Ranges.forRange("Test!A4:C40").clearStartRow().clearEndRow().toRange();
      fail();
    } catch (IllegalStateException e) {
      // expected exception from not clearing coordinates out of order.
    }
  }

  @Test
  public void forRange_expandRowsColumnsRange() {
    assertEquals("Test!A13:F20",
        Ranges.forRange("Test!A13:C15").expandColumns(3).expandRows(5).toRange());
  }

  @Test
  public void forRange_expandColumnsNoEndColumn() {
    try {
      Ranges.forRange("'Brian''s Sheet'!A1:5").expandColumns(4).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected: Cannot expand a range columns where the end column isn't set.
    }
  }

  @Test
  public void forRange_expandRowsNoStartColumn() {
    try {
      Ranges.forRange("'Brian''s Sheet'").expandColumns(4).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected: Cannot expand a range rows where the start column isn't set.
    }
  }

  @Test
  public void forRange_expandColumnsLessThanOne() {
    try {
      Ranges.forRange("'Brian''s Sheet'!A1:B5").expandColumns(0).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Expected: Cannot expand by columns < 1
    }
  }

  @Test
  public void forRange_expandRowsNoEndRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'!A1:C").expandRows(4).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected: Cannot expand a range rows where the end row isn't set.
    }
  }

  @Test
  public void forRange_expandRowsNoStartRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'").expandRows(4).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected: Cannot expand a range rows where the start row isn't set.
    }
  }

  @Test
  public void forRange_expandRowsLessThanOne() {
    try {
      Ranges.forRange("'Brian''s Sheet'!A1:B5").expandRows(0).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Expected: Cannot expand by rows < 1
    }
  }

  @Test
  public void forRange_translateNoStartColumnBound() {
    try {
      Ranges.forRange("'Brian''s Sheet'").translate(1, 0).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected: Cannot translate where start column bound is not set.
    }
  }

  @Test
  public void forRange_translateNoStartRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'").translate(0, 1).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected: Cannot translate where start row is not set.
    }
  }

  @Test
  public void forRange_translateBeyondStartColumnBound() {
    try {
      Ranges.forRange("'Brian''s Sheet'!C5").translate(-4, 0).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      System.out.println(e.getMessage());
      // Expected: Cannot translate beyond the start column
    }
  }

  @Test
  public void forRange_translateBeyondStartRow() {
    try {
      Ranges.forRange("'Brian''s Sheet'!C5").translate(0, -5).toRange();
      fail();
    } catch (IllegalArgumentException e) {
      // Expected: Cannot translate to above the start row.
    }
  }

  @Test
  public void forRange_translateRange() {
    assertEquals("Test!D6:G11",
        Ranges.forRange("Test!A1:D6").translate(3, 5).toRange());
  }

  @Test
  public void forRange_translateRangeColumnOnly() {
    assertEquals("Test!D1:G6",
        Ranges.forRange("Test!A1:D6").translate(3, 0).toRange());
  }

  @Test
  public void forRange_translateRangeCellOnly() {
    assertEquals("Test!D6",
        Ranges.forRange("Test!A1").translate(3, 5).toRange());
  }

  @Test
  public void forRange_rowAndColumnOnlyPair() {
    try {
      Ranges.forRange("'Brian''s Sheet'!15:Z");
      fail();
    } catch (IllegalArgumentException e) {
      // Expected exception for range of form ...!<row>:<column> or ...!<column>:<row>
    }
  }

  @Test
  public void forSheetName_toValidUnboundedRamge() {
    assertEquals("'Today''s Metrics'", Ranges.forSheetName("Today's Metrics").toRange());
  }

  @Test
  public void forSheetName_toValidBoundedRamge() {
    assertEquals("'Today''s Metrics'!B3:C5",
            Ranges.forSheetName("Today's Metrics")
                    .withStartColumn(1)
                    .withStartRow(2)
                    .withEndColumn(2)
                    .withEndRow(4)
                    .toRange());
  }

  @Test
  public void forSheetName_endColumnWithNoStartColumn() {
    try {
      Ranges.forSheetName("Today's Metrics")
          .withEndColumn(2)
          .toRange();
      fail();
    } catch (IllegalStateException e) {
      // expected exception for setting an end bound without a start bound.
    }
  }

  @Test
  public void forSheetName_endRowWithNoStartRow() {
    try {
      Ranges.forSheetName("Today's Metrics")
          .withEndRow(4)
          .toRange();
      fail();
    } catch (IllegalStateException e) {
      // expected exception for setting an end bound without a start bound.
    }
  }

  @Test
  public void forSheetName_toValidSingleCellRange() {
    assertEquals("'Today''s Metrics'!B3",
            Ranges.forSheetName("Today's Metrics")
                    .withStartColumn(1)
                    .withStartRow(2)
                    .withEndColumn(1)
                    .withEndRow(2)
                    .toRange());
  }

  @Test
  public void forGridRange() {
    GridRange gridRange = new GridRange();
    gridRange.setSheetId(0);
    gridRange.setStartColumnIndex(0);
    gridRange.setStartRowIndex(0);
    gridRange.setEndColumnIndex(10);
    gridRange.setEndRowIndex(10);
    assertEquals("Test!A1:J10",
        Ranges.forGridRange(gridRange).withSheetName("Test").toRange());
  }

  @Test
  public void forGridRange_noSheetNameSet() {
    try {
      GridRange gridRange = new GridRange();
      gridRange.setSheetId(0);
      gridRange.setStartColumnIndex(0);
      gridRange.setStartRowIndex(0);
      gridRange.setEndColumnIndex(10);
      gridRange.setEndRowIndex(10);
      Ranges.forGridRange(gridRange).toRange();
      fail();
    } catch (IllegalStateException e) {
      // Expected, cannot convert to a range where SheetName is not set.
    }
  }

  @Test
  public void forGridRange_startGridCoordinate() {
    GridRange gridRange = new GridRange();
    gridRange.setSheetId(0);
    gridRange.setStartColumnIndex(0);
    gridRange.setStartRowIndex(0);
    gridRange.setEndColumnIndex(10);
    gridRange.setEndRowIndex(10);
    GridCoordinate start = Ranges
        .forGridRange(gridRange).withSheetName("Test").toStartGridCoordinate();
    assertEquals(0, start.getColumnIndex().intValue());
    assertEquals(0, start.getRowIndex().intValue());
  }

  @Test
  public void forGridRange_endGridCoordinate() {
    GridRange gridRange = new GridRange();
    gridRange.setSheetId(0);
    gridRange.setStartColumnIndex(0);
    gridRange.setStartRowIndex(0);
    gridRange.setEndColumnIndex(10);
    gridRange.setEndRowIndex(10);
    GridCoordinate start = Ranges
        .forGridRange(gridRange).withSheetName("Test").toEndGridCoordinate();
    assertEquals(9, start.getColumnIndex().intValue());
    assertEquals(9, start.getRowIndex().intValue());
  }

  @Test
  public void forStartGridCoordinate() {
    GridCoordinate gridCoordinate = new GridCoordinate();
    gridCoordinate.setSheetId(0);
    gridCoordinate.setColumnIndex(0);
    gridCoordinate.setRowIndex(0);
    assertEquals("Test!A1:J10",
        Ranges.forStartGridCoordinate(gridCoordinate)
            .withSheetName("Test").withWidth(10).withHeight(10).toRange());
  }

  @Test
  public void forSheet() {
    SheetProperties properties = new SheetProperties();
    properties.setTitle("Today's results!");
    properties.setSheetId(123);
    Sheet sheet = new Sheet();
    sheet.setProperties(properties);
    assertEquals("'Today''s results!'", Ranges.forSheet(sheet).toRange());
  }
}