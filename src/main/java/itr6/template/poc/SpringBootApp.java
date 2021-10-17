package itr6.template.poc;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

@SpringBootApplication
@Slf4j
public class SpringBootApp implements CommandLineRunner {
  private static final String GENERAL_DATA_SHEET_NAME = "PARTAGENERAL";
  private static final String ITR6_GENERAL_SHEET_NAME = "PART A - GENERAL";

  public static void main(String[] args) {
    SpringApplication.run(SpringBootApp.class, args);
  }

  @Override
  public void run(String... args) throws Exception {
    log.info("Reading configuration source...");
    //Read Configuration file
    Map<String, String> indexConfigurations = new HashMap<>();
    Workbook configSource = WorkbookFactory.create(new File("/Users/s1b06wv/Downloads/ConfigScource.xlsx"));
    Sheet generalConfigSheet = configSource.getSheet(GENERAL_DATA_SHEET_NAME);
    for(Row row: generalConfigSheet){
      if(row.getRowNum() == 0){
        continue;//skip header
      }
      Cell fieldNameCell = row.getCell(0);
      Cell indexCell = row.getCell(1);
      if(fieldNameCell != null && indexCell != null){
        indexConfigurations.put(fieldNameCell.getStringCellValue(), indexCell.getStringCellValue());
      }
    }
    log.info("Completed reading configuration source");
    log.info("Reading data source...");
    //Read data source
    Map<String, String> dataMap = new HashMap<>();
    Workbook dataSource = WorkbookFactory.create(new File("/Users/s1b06wv/Downloads/DataSource.xlsx"));
    Sheet generalDataSheet = dataSource.getSheet(GENERAL_DATA_SHEET_NAME);
    for(Row row: generalDataSheet){
      if(row.getRowNum() == 0){
        continue;//skip header
      }
      Cell fieldNameCell = row.getCell(0);
      Cell valueCell = row.getCell(1);
      if(fieldNameCell != null && valueCell != null){
        dataMap.put(fieldNameCell.getStringCellValue(), valueCell.getStringCellValue());
      }
    }
    log.info("Completed reading data source");
    log.info("Filling ITR6 Template...");
    //Write data to template
    Workbook itr6Template = WorkbookFactory.create(new File("/Users/s1b06wv/Downloads/ITR6_V1.0.xlsm"));
    Sheet itr6GeneralSheet = itr6Template.getSheet(ITR6_GENERAL_SHEET_NAME);
    indexConfigurations.forEach((field_name, cellIndex) -> {
      CellReference ref = new CellReference(cellIndex);
      Row r = itr6GeneralSheet.getRow(ref.getRow());
      if (r != null) {
        Cell c = r.getCell(ref.getCol());
        if(c!= null)
          c.setCellValue(dataMap.get(field_name));
      }
    });
    FileOutputStream os = new FileOutputStream("/Users/s1b06wv/Downloads/Filled_ITR6_V1.0.xlsm");
    itr6Template.write(os);
    itr6Template.close();
    os.close();
    log.info("Completed ITR6 template filling");
  }
}
