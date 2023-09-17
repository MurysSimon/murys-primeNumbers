package murys.primeNumbers;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelManager {

    public ExcelManager() {
    }

    public void getExcelPath(){
        // ziskani cesty k souboru od uzivatele
        System.out.println("Zadejte ABSOLUTNI CESTU k Vasemu excel souboru (xlsx):");
        Scanner scanner = new Scanner(System.in);
        String excelPath = scanner.nextLine();
        scanner.close();

        try{
            File file = new File(excelPath);
            // zjisteni zdali cesta obsahuje priponu xlsx, pokud ne tak ji prida automaticky
            if(!FilenameUtils.getExtension(excelPath).equals("xlsx")) {
                excelPath += ".xlsx";
                file = new File(excelPath);
            }
            // zde se podminka pta jestli soubor existuje, jestli je citelny a jestli se vubec jedna o soubor
            if(file.exists() && file.canRead() && file.isFile()) {
                excelRead(excelPath);
            }
            else {
                System.out.println("Zadana cesta k excel (xlsx) souboru neexistuje a nebo soubor neni citelny!");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /* pro praci se soubory excel jsem se rozhodl pouzit knihovnu apache poi, jelikoz se zda byt nejrozsirenejsi
        a nejpouzivanejsi pro praci s excel soubory v jave
     */
    private void excelRead(String givenPath){
        try {
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(new FileInputStream(givenPath));

            // prochazeni jednotlivych radku a bunek souboru
            // prvni smycka resi jednotlive listy souboru excel
            for(int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(i);

                // tyto dve smycky prochazi samotny list radek po radku, bunku po bunce
                for(Row row : xssfSheet){
                    for(Cell cell : row){
                        if(cell.getCellType() == CellType.NUMERIC) {    // podminka pro kontrolu typu hodnoty bunky
                            int cellNumber = (int) cell.getNumericCellValue();
                            if(isNumberPrime(cellNumber)) {
                                System.out.print(cell + "\n");
                            }
                        }
                        if(cell.getCellType() == CellType.STRING){  // jelikoz jsou data v poskytnutem excelu ve formatu "text", rozhodl jsem se testovat i bunky tohoto typu
                            String cellValue = cell.getStringCellValue();
                            try{
                                int cellNumber = Integer.parseInt(cellValue);
                                if(isNumberPrime(cellNumber)) {
                                    System.out.print(cellValue + "\n");
                                }
                            } catch (NumberFormatException e) {};
                        }
                    }
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static boolean isNumberPrime(int number) {
        if (number < 2) return false; // pokud je mensi nez 2 tak nemuze byt prvocislo
        for (int i = 2; i <= Math.sqrt(number); i++) {
            if (number % i == 0) return false; // kontrola zda number je delitelne "i" beze zbytku, pokud je tak nemuze byt prvocislo
        }
        return true;
    }
}
