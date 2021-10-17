package itr6.template.poc;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

@SpringBootApplication
@Slf4j
public class SpringBootApp implements CommandLineRunner {
  private static final String[] SHEETS_TO_READ = new String[]{"PART A - GENERAL", "GENERAL2", "NATURE OF BUSINESS"};

  @Value("${config.location}")
  private String configFile;

  @Value("${data.location}")
  private String dataFile;

  @Value("${template.location}")
  private String templateFile;

  @Value("${filled.template.out.location}")
  private String outputFile;

  public static void main(String[] args) {
    SpringApplication.run(SpringBootApp.class, args);
  }

  @Override
  public void run(String... args) throws Exception {
    log.info("STARTED ITR6 template filling");

    Workbook itr6Template = WorkbookFactory.create(new FileInputStream(templateFile));
    Workbook configSource = WorkbookFactory.create(new FileInputStream(configFile));
    Workbook dataSource = WorkbookFactory.create(new FileInputStream(dataFile));

    for (String sheetName : SHEETS_TO_READ) {
      log.info("Sheet name : {}, READING configuration", sheetName);
      //Read Configuration file
      Map<String, String> indexConfigurations = new HashMap<>();
      Sheet configSheet = configSource.getSheet(sheetName);
      for (Row row : configSheet) {
        if (row.getRowNum() == 0) {
          continue;//skip header
        }
        Cell fieldNameCell = row.getCell(0);
        Cell indexCell = row.getCell(1);
        if (fieldNameCell != null && indexCell != null && StringUtils.isNotBlank(fieldNameCell.getStringCellValue())) {
          indexConfigurations.put(fieldNameCell.getStringCellValue(), indexCell.getStringCellValue());
        }
      }
      log.info("Sheet name : {}, COMPLETED reading configuration", sheetName);
      log.info("Sheet name : {}, READING data", sheetName);
      //Read data source
      Map<String, String> dataMap = new HashMap<>();
      Sheet dataSheet = dataSource.getSheet(sheetName);
      for (Row row : dataSheet) {
        if (row.getRowNum() == 0) {
          continue;//skip header
        }
        Cell fieldNameCell = row.getCell(0);
        Cell valueCell = row.getCell(1);
        if (fieldNameCell != null && valueCell != null && StringUtils.isNotBlank(fieldNameCell.getStringCellValue())) {
          dataMap.put(fieldNameCell.getStringCellValue(), valueCell.getStringCellValue());
        }
      }
      log.info("Sheet name : {}, COMPLETED reading data", sheetName);
      log.info("Sheet name : {}, STARTED filling ITR6 template", sheetName);
      //Write data to template
      Sheet itr6TemplateSheet = itr6Template.getSheet(sheetName);
      indexConfigurations.forEach((field_name, cellIndex) -> {
        CellReference ref = new CellReference(cellIndex);
        Row r = itr6TemplateSheet.getRow(ref.getRow());
        if (r != null) {
          Cell c = r.getCell(ref.getCol());
          if (c != null)
            c.setCellValue(dataMap.get(field_name));
        }
      });
      log.info("Sheet name : {}, COMPLETED filling ITR6 template", sheetName);
    }
    log.info("SAVING ITR6 template...");
    FileOutputStream os = new FileOutputStream(outputFile);
    itr6Template.write(os);
    itr6Template.close();
    configSource.close();
    dataSource.close();
    os.close();
    log.info("COMPLETED ITR6 template filling");
  }
}
