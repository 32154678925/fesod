package org.apache.fesod.sheet.writesheet;

import org.apache.fesod.sheet.ExcelWriter;
import org.apache.fesod.sheet.FesodSheet;
import org.apache.fesod.sheet.annotation.ExcelProperty;
import org.apache.fesod.sheet.support.ExcelTypeEnum;
import org.apache.fesod.sheet.util.TestFileUtil;
import org.apache.fesod.sheet.write.builder.ExcelWriterBuilder;
import org.apache.fesod.sheet.write.builder.ExcelWriterSheetBuilder;
import org.apache.fesod.sheet.write.metadata.WriteSheet;
import org.apache.fesod.sheet.write.metadata.WriteTable;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.File;
import java.lang.reflect.Field;
import java.util.*;

public class WriteTableTest {

    private static final String HEADER_STRING;

    static {
        try {
            Field field = WriteSheetData.class.getDeclaredField("string");
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            HEADER_STRING = annotation.value()[0];
        } catch (Exception e) {
            throw new RuntimeException(
                    "Failed to resolve @ExcelProperty header from WriteSheetData.string", e);
        }
    }

    @ParameterizedTest
    @EnumSource(value = ExcelTypeEnum.class, names = {"XLS", "XLSX"})
    public void testWriteTableWithTitle(ExcelTypeEnum excelType) throws Exception {
        Random random = new Random();
        int offsetA = random.nextInt(10);
        int offsetB = random.nextInt(10);
        int sizeA = random.nextInt(10)+1;
        int sizeB = random.nextInt(10)+1;

        File testFile = TestFileUtil.createNewFile(
                "writesheet/write-table-title" + excelType.getValue());

        // ── write ────────────────────────────────────────────────────────────

        ExcelWriterBuilder write = FesodSheet.write(testFile);
        ExcelWriterSheetBuilder sheetBuilder = write.sheet(0, "0");
        WriteSheet sheet = sheetBuilder.build();

        WriteTable tableA = sheetBuilder.table(0)
                .head(WriteSheetData.class)
                .relativeHeadRowIndex(offsetA)
                .build();
        WriteTable tableB = sheetBuilder.table(1)
                .relativeHeadRowIndex(offsetB)
                .head(WriteSheetData.class)
                .build();

        try (ExcelWriter writer = write.build()) {
            writer.write(getList(sizeA, "A-"), sheet, tableA);
            writer.write(getList(sizeB, "B-"), sheet, tableB);
            writer.finish();
        }

        Assertions.assertTrue(testFile.exists(), "Written file should exist");
        Assertions.assertTrue(testFile.length() > 0, "Written file should not be empty");

        // ── read & assert ────────────────────────────────────────────────────

        try (Workbook workbook = WorkbookFactory.create(testFile)) {
            Sheet excelSheet = workbook.getSheetAt(0);
            Assertions.assertNotNull(excelSheet, "Sheet '0' should exist");

            Map<Integer, String> expected = buildExpectedWithTitle(offsetA, offsetB, sizeA, sizeB);
            int totalRows = expected.size();
            Assertions.assertEquals(totalRows, excelSheet.getPhysicalNumberOfRows(),
                    "Sheet should have exactly " + totalRows + " rows "
                            + "(offsets: A=" + offsetA + " B=" + offsetB
                            + ", sizes: A=" + sizeA + " B=" + sizeB + ")");

            for (Map.Entry<Integer, String> entry : expected.entrySet()) {
                int rowIdx = entry.getKey();
                String expectedValue = entry.getValue();
                Row row = excelSheet.getRow(rowIdx);
                Assertions.assertNotNull(row,
                        "Expected row " + rowIdx + " with value '" + expectedValue
                                + "' (A:" + offsetA + "/" + sizeA
                                + " B:" + offsetB + "/" + sizeB + ")");
                Cell cell = row.getCell(0);
                Assertions.assertNotNull(cell,
                        "Expected cell at row " + rowIdx);
                Assertions.assertEquals(expectedValue, cell.getStringCellValue(),
                        "Row " + rowIdx + " expected '" + expectedValue + "' "
                                + "(A:" + offsetA + "/" + sizeA
                                + " B:" + offsetB + "/" + sizeB + ")");
            }
        }
    }

    @ParameterizedTest
    @EnumSource(value = ExcelTypeEnum.class, names = {"XLS", "XLSX"})
    public void testWriteTableWithoutTitle(ExcelTypeEnum excelType) throws Exception {
        Random random = new Random();
        int offsetC = random.nextInt(10);
        int offsetA = random.nextInt(10);
        int offsetB = random.nextInt(10);
        int sizeC = random.nextInt(10)+1;
        int sizeA = random.nextInt(10)+1;
        int sizeB = random.nextInt(10)+1;

        File testFile = TestFileUtil.createNewFile(
                "writesheet/write-table-notitle" + excelType.getValue());

        // ── write ────────────────────────────────────────────────────────────

        ExcelWriterBuilder write = FesodSheet.write(testFile);
        ExcelWriterSheetBuilder sheetBuilder = write.sheet(0, "0");
        WriteSheet sheet = sheetBuilder.build();

        WriteSheet collect = sheetBuilder
                .relativeHeadRowIndex(offsetC)
                .build();
        WriteTable tableA = sheetBuilder.table(0)
                .relativeHeadRowIndex(offsetA)
                .build();
        WriteTable tableB = sheetBuilder.table(1)
                .relativeHeadRowIndex(offsetB)
                .build();

        try (ExcelWriter writer = write.build()) {
            writer.write(getList(sizeC, "C-"), collect);
            writer.write(getList(sizeA, "A-"), sheet, tableA);
            writer.write(getList(sizeB, "B-"), sheet, tableB);
            writer.finish();
        }

        Assertions.assertTrue(testFile.exists(), "Written file should exist");
        Assertions.assertTrue(testFile.length() > 0, "Written file should not be empty");

        // ── read & assert ────────────────────────────────────────────────────

        try (Workbook workbook = WorkbookFactory.create(testFile)) {
            Sheet excelSheet = workbook.getSheetAt(0);
            Assertions.assertNotNull(excelSheet, "Sheet '0' should exist");

            Map<Integer, String> expected = buildExpectedWithoutTitle(
                    offsetC, offsetA, offsetB, sizeC, sizeA, sizeB);
            int totalRows = expected.size();
            Assertions.assertEquals(totalRows, excelSheet.getPhysicalNumberOfRows(),
                    "Sheet should have exactly " + totalRows + " rows "
                            + "(C:" + offsetC + "/" + sizeC
                            + " A:" + offsetA + "/" + sizeA
                            + " B:" + offsetB + "/" + sizeB + ")");

            for (Map.Entry<Integer, String> entry : expected.entrySet()) {
                int rowIdx = entry.getKey();
                String expectedValue = entry.getValue();
                Row row = excelSheet.getRow(rowIdx);
                Assertions.assertNotNull(row,
                        "Expected row " + rowIdx + " with value '" + expectedValue
                                + "' (C:" + offsetC + "/" + sizeC
                                + " A:" + offsetA + "/" + sizeA
                                + " B:" + offsetB + "/" + sizeB + ")");
                Cell cell = row.getCell(0);
                Assertions.assertNotNull(cell,
                        "Expected cell at row " + rowIdx);
                Assertions.assertEquals(expectedValue, cell.getStringCellValue(),
                        "Row " + rowIdx + " expected '" + expectedValue + "' "
                                + "(C:" + offsetC + "/" + sizeC
                                + " A:" + offsetA + "/" + sizeA
                                + " B:" + offsetB + "/" + sizeB + ")");
            }
        }
    }

    // ── expected-row builders ────────────────────────────────────────────────

    /**
     * Build expected row map for with-title tables.
     * Tables are sequential: B starts after A's last row + offsetB empty rows.
     */
    private static Map<Integer, String> buildExpectedWithTitle(
            int offsetA, int offsetB, int sizeA, int sizeB) {
        Map<Integer, String> expected = new LinkedHashMap<>();

        // Table A: head at offsetA, data at offsetA+1 .. offsetA+sizeA
        expected.put(offsetA, HEADER_STRING);
        for (int i = 0; i < sizeA; i++) {
            expected.put(offsetA + 1 + i, "A-" + i);
        }

        // Table B: offsetB empty rows after A's last row, then head + data
        int bHeadRow = offsetA + sizeA + offsetB + 1;
        expected.put(bHeadRow, HEADER_STRING);
        for (int i = 0; i < sizeB; i++) {
            expected.put(bHeadRow + 1 + i, "B-" + i);
        }

        return expected;
    }

    /**
     * Build expected row map for headerless writes.
     * Tables are sequential: each starts after the previous one's last row + offset.
     */
    private static Map<Integer, String> buildExpectedWithoutTitle(
            int offsetC, int offsetA, int offsetB, int sizeC, int sizeA, int sizeB) {
        Map<Integer, String> expected = new LinkedHashMap<>();

        // collect: data at offsetC .. offsetC+sizeC-1
        for (int i = 0; i < sizeC; i++) {
            expected.put(offsetC + i, "C-" + i);
        }

        // tableA: starts after collect's last row + offsetA empty rows
        int aStartRow = offsetC + sizeC + offsetA;
        for (int i = 0; i < sizeA; i++) {
            expected.put(aStartRow + i, "A-" + i);
        }

        // tableB: starts after A's last row + offsetB empty rows
        int bStartRow = aStartRow + sizeA + offsetB;
        for (int i = 0; i < sizeB; i++) {
            expected.put(bStartRow + i, "B-" + i);
        }

        return expected;
    }

    // ── helper ───────────────────────────────────────────────────────────────

    private static List<WriteSheetData> getList(int size, String prefix) {
        List<WriteSheetData> dataList = new ArrayList<>();
        for (int j = 0; j < size; j++) {
            WriteSheetData data = new WriteSheetData();
            data.setString(prefix + j);
            dataList.add(data);
        }
        return dataList;
    }
}
