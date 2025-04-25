package tpm.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelReader {
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("^\\d+");

    private ExcelReader() throws IOException, InvalidFormatException {
    }

    public static void main(String[] args) throws IOException, InvalidFormatException {
        File excelFile = Paths.get("C:\\Users\\user\\IdeaProjects\\ExcelReader\\src\\main\\resources")
                .resolve("excel.xlsx")
                .toFile();
        ExcelReader reader = new ExcelReader();
        reader.readFromExcelFile(excelFile);
    }

    public void readFromExcelFile(File excelFile) throws IOException {
        List<ItemField> itemFields = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(excelFile)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {

                if (row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) != null) {

                    if (row.getCell(0).getCellType() == CellType.STRING) {
                        Matcher matcher = PLACEHOLDER_PATTERN.matcher(row.getCell(0).getStringCellValue());
                        String cellValue = row.getCell(0).getStringCellValue();
                        if (!matcher.find()) {
                            continue;
                        }
                    }

                    try {
                        itemFields.add(returnRowAsField(row));
                    } catch (RuntimeException e) {
                        System.out.println(e);
                    }
                }
            }
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        } catch (RuntimeException e) {
            System.out.println(e);
        }

        itemFields.stream().limit(10)
                .forEach((i) -> System.out.println(i.getBarcode()));
    }

    public ItemField returnRowAsField(Row row) {
        ItemField fields = new ItemField();
        for (int i = 0; i < 9; i++) {
            if (row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL) == null) {
                continue;
            }
            switch (i) {
                case 0:
                    if (row.getCell(i).getCellType() == CellType.STRING) {
                        fields.setBarcode(Long.parseLong(row.getCell(i).getStringCellValue()));
                    }
                    else {
                        fields.setBarcode((long) row.getCell(i).getNumericCellValue());
                    }
                    break;
                case 1:
                    fields.setProductName(row.getCell(i).getStringCellValue());
                    break;
                case 2:
                    fields.setDescription(row.getCell(i).getStringCellValue());
                    break;
                case 3:
                    fields.setNutrition(row.getCell(i).getStringCellValue());
                    break;
                case 4:
                    fields.setBrand(row.getCell(i).getStringCellValue());
                    break;
                case 5:
                    if (row.getCell(i).getCellType() == CellType.STRING) {
                        fields.setExpiration(Integer.parseInt(row.getCell(i).getStringCellValue()));
                    }else {
                        fields.setExpiration((int) row.getCell(i).getNumericCellValue());
                    }
                    break;
                case 6:
                    fields.setStorageCondition(row.getCell(i).getStringCellValue());
                    break;
                case 7:
                    fields.setComposition(row.getCell(i).getStringCellValue());
                    break;
                case 8:
                    fields.setImage(row.getCell(i).getStringCellValue());
                    break;
                default:
                    break;
            }
        }
        return fields;
    }

}
