package org.apache.fesod.sheet.writesheet;

import org.apache.fesod.sheet.ExcelWriter;
import org.apache.fesod.sheet.FesodSheet;
import org.apache.fesod.sheet.support.ExcelTypeEnum;
import org.apache.fesod.sheet.util.TestFileUtil;
import org.apache.fesod.sheet.write.builder.ExcelWriterBuilder;
import org.apache.fesod.sheet.write.builder.ExcelWriterSheetBuilder;
import org.apache.fesod.sheet.write.metadata.WriteSheet;
import org.apache.fesod.sheet.write.metadata.WriteTable;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.*;

public class WriteTableTest {
    
    @Test
    public void testWriteTableWithTitle(){
        File testFile = TestFileUtil.createNewFile("writesheet/write-table" + ExcelTypeEnum.XLS.getValue());
        ExcelWriterBuilder write = FesodSheet.write(testFile);
        
        ExcelWriterSheetBuilder sheetBuilder = write.sheet(0, "0");
        WriteSheet sheet = sheetBuilder.build();
        
        WriteTable tableA = sheetBuilder.table(0)
                .head(WriteSheetData.class)
                .relativeHeadRowIndex(3)
                .build();
        
        WriteTable tableB = sheetBuilder.table(1)
                .relativeHeadRowIndex(6)
                .head(WriteSheetData.class)
                .build();
        
        try (ExcelWriter writer = write.build()) {
            writer.write(getList(), sheet, tableA);
            
            writer.write(getList(), sheet, tableB);
            
            writer.finish();
        }
    }
    
    @Test
    public void testWriteTableWithoutTitle() {
        File testFile = TestFileUtil.createNewFile("writesheet/write-table" + ExcelTypeEnum.XLS.getValue());
        ExcelWriterBuilder write = FesodSheet.write(testFile);
        ExcelWriterSheetBuilder sheetBuilder = write.sheet(0, "0");
        WriteSheet sheet = sheetBuilder.build();
        
        WriteSheet collect = sheetBuilder
                .relativeHeadRowIndex(1)
                .build();
        
        WriteTable tableA = sheetBuilder.table(0)
                .relativeHeadRowIndex(3)
                .build();
        
        WriteTable tableB = sheetBuilder.table(1)
                .relativeHeadRowIndex(6)
                .build();
        
        try (ExcelWriter writer = write.build()) {
            writer.write(getList(), collect);
            writer.write(getList(), sheet, tableA);
            writer.write(getList(), sheet, tableB);
            writer.finish();
        }
    }
    
    private List<WriteSheetData> getList(){
        Random random = new Random();
        List<WriteSheetData> dataList = new ArrayList<>();
        for (int j = 0; j < 10; j++) {
            WriteSheetData testA = new WriteSheetData();
            testA.setString(random.nextInt(100) + "");
            dataList.add(testA);
        }
        return dataList;
    }
}
