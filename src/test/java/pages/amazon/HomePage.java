package pages.amazon;

import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindAll;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import utilities.ConfigReader;
import utilities.Driver;
import utilities.SeleniumUtils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class HomePage {
    public HomePage(){
        PageFactory.initElements(Driver.getDriver(), this);
    }

    private String product = ConfigReader.getProperty("product");
    private String filePath = ConfigReader.getProperty("file_path");

    private String productNameXpath = "//span[contains(@cel_widget_id, 'MAIN-SEARCH_RESULTS-')]/div/div//h2//a//span";
    private String priceSymbolXpath = "//div[@data-component-type='s-search-result']//span[@class='a-price-symbol']";

    @FindBy(id = "twotabsearchtextbox")
    WebElement searchInput;

    @FindAll(
            @FindBy(xpath = "//span[contains(@cel_widget_id, 'MAIN-SEARCH_RESULTS-')]/div/div")
    ) List<WebElement> searchResults;

    @FindAll(
            @FindBy(xpath = "//span[contains(@cel_widget_id, 'MAIN-SEARCH_RESULTS-')]/div/div//h2//a//span")
    ) List<WebElement> productNames;

    @FindAll(
            @FindBy(xpath = "//div[@data-component-type='s-search-result']//span[@class='a-price-symbol']")
    ) List<WebElement> priceSymbols;

    @FindAll(
            @FindBy(xpath = "//div[@data-component-type='s-search-result']//span[@class='a-price-whole']")
    ) List<WebElement> prices;

    @FindBy(xpath = "//ul[@class='a-pagination']//li[a[contains(text(), 'Volgende')]]")
    WebElement nextBtn;


    public void searchProduct(){
        SeleniumUtils.waitForVisibility(searchInput, 5);
        searchInput.clear();
        searchInput.sendKeys(product + Keys.ENTER);
    }

    private List<Map<String,String>> getProductInfos(){
        List<Map<String,String>> ls = new ArrayList<>();
        Map<String,String> map = new HashMap<>();

        String name, symbol, price;
        String notExist = "not exist";

        for(int i = 0; i<searchResults.size(); i++){
            if(SeleniumUtils.isElementExist(searchResults.get(i), productNameXpath)){
                name = searchResults.get(i).findElements(By.xpath(productNameXpath)).get(i).getText();
            }else name = notExist;
            if(SeleniumUtils.isElementExist(searchResults.get(i), priceSymbolXpath)){
                symbol = searchResults.get(i).findElements(By.xpath(priceSymbolXpath)).get(i).getText();
            }else symbol = notExist;
            if(SeleniumUtils.isElementExist(searchResults.get(i), priceSymbolXpath)){
                price = searchResults.get(i).findElements(By.xpath(priceSymbolXpath)).get(i).getText();
            }else price = notExist;

            map.put("name", name);
            map.put("symbol", symbol);
            map.put("price", price);

            ls.add(map);
        }
        return ls;
    }

    public void writeToExcel(){
        List<Map<String,String>> ls = getProductInfos();
        boolean isDisabled = false;
        int i = 0;

            try{
                //Open the file/workbook
                FileInputStream fileInputStream = new FileInputStream(filePath);
                Workbook        workbook        = WorkbookFactory.create(fileInputStream);
                //Open the first worksheet
                Sheet sheet =workbook.getSheetAt(0);
                Row row;
                //write headers
                row = sheet.getRow(0);
                row.createCell(0).setCellValue("name");
                row.createCell(1).setCellValue("symbol");
                row.createCell(2).setCellValue("price");

                while(!isDisabled){
                    //Go to the second row
                    row = sheet.getRow(i+1);
                    row.createCell(0).setCellValue(ls.get(i).get("name"));
                    row.createCell(1).setCellValue(ls.get(i).get("symbol"));
                    row.createCell(2).setCellValue(ls.get(i).get("price"));

                    isDisabled = nextBtn.getAttribute("class").contains("a-disabled");
                    nextBtn.click();
                    i++;
                }
                //Write and save the workbook
                //FileInputStream is to READ, FileOutputStream is to WRITE
                FileOutputStream fileOutputStream = new FileOutputStream(filePath);
                workbook.write(fileOutputStream);

                //Close the file
                fileInputStream.close();
                fileOutputStream.close();
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
    }

}
