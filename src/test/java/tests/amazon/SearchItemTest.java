package tests.amazon;

import org.testng.annotations.Test;
import pages.amazon.HomePage;

public class SearchItemTest extends TestBase{
    HomePage homePage;
    @Test
    public void writeToExcel(){
        homePage = new HomePage();

        homePage.searchProduct();
        homePage.writeToExcel();
    }
}
