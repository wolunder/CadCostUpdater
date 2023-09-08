package parsers;

import model.CadNumber;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.*;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import javax.swing.*;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;

public class FondCadCost implements ParserImpl {
    private WebElement elementCost;
    private WebElement elementFindField;
    private WebElement elementSort;
    private WebElement elementData;
    private static final String FOND_CAD_COST_URL = "https://rosreestr.gov.ru/wps/portal/cc_ib_svedFDGKO";
    private static final String FIND_FIELD_XPATH = "//*[@id=\"fdgko_search_string\"]";
    private static final String COST_XPATH = "/html/body/div[1]/div[6]/div[4]/div/div[1]/section/div[2]/div[2]" +
            "/div[3]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[2]";
    private static final String DATA_XPATH = "/html/body/div[1]/div[6]/div[4]/div/div[1]/section/div[2]/div[2]" +
            "/div[3]/div[2]/div[2]/div/div/table/tbody/tr[1]/td[1]";
    private static final String SORT_XPATH = "/html/body/div[1]/div[6]/div[4]/div/div[1]/section/div[2]/div[2]" +
            "/div[3]/div[2]/div[2]/div/div/table/thead/tr[1]/th[1]/a";
    private static final String DATA_SECOND_XPATH = "/html/body/div[1]/div[6]/div[4]/div/div[1]/section/div[2]/div[2]" +
            "/div[3]/div[2]/div[2]/div/div/table/tbody/tr[2]/td[1]";
    private static final String COST_SECOND_XPATH = "/html/body/div[1]/div[6]/div[4]/div/div[1]/section/div[2]" +
            "/div[2]/div[3]/div[2]/div[2]/div/div/table/tbody/tr[2]/td[2]/a";

    @Override
    public List<CadNumber> getListCad(Set<String> cadNumSet, WebDriver driver) {
        driver.get(FOND_CAD_COST_URL);
        driver.manage().window().minimize();
        driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));
        List<CadNumber> cadNumberList = new ArrayList<>();

        for (String cadNum : cadNumSet) {
            CadNumber cad = new CadNumber();
            cad.setCadNumber(cadNum);
            try {
                driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));
                //окно поиска вбиваем кадастровый
                elementFindField = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                        .until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath(FIND_FIELD_XPATH)));
                elementFindField.sendKeys(cadNum);
                elementFindField.sendKeys(Keys.RETURN);
                driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));

                elementSort = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                        .until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath(SORT_XPATH)));
                elementSort.click();//сортируем
                driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));

                elementSort = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                        .until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath(SORT_XPATH)));
                elementSort.click();
                driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(60));
                //ищем стоимость
                elementCost = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                        .until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath(COST_XPATH)));
                elementData = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                        .until(ExpectedConditions.presenceOfElementLocated(
                                By.xpath(DATA_XPATH)));
                if (elementData.getText().contains("-")) {
                    elementCost = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                            .until(ExpectedConditions.presenceOfElementLocated(
                                    By.xpath(COST_SECOND_XPATH)));
                    elementData = (new WebDriverWait(driver, Duration.ofSeconds(60)))
                            .until(ExpectedConditions.presenceOfElementLocated(
                                    By.xpath(DATA_SECOND_XPATH)));
                }

                cad.setData(elementData.getText());
                if(elementCost.getText().matches("^[0-9]*[.,][0-9]+$")) {
                    cad.setCost(Double.parseDouble(elementCost.getText().replace(",", ".")));
                }else {
                    cad.setCost(0);
                }
                if(elementCost != null && elementData != null) {
                    System.out.println(elementCost.getText() + " ---- " + elementData.getText());
                }

            } catch (TimeoutException | NoSuchElementException | StaleElementReferenceException e) {
                cad.setCost(0);
            } finally {
                driver.navigate().to(FOND_CAD_COST_URL);
                cadNumberList.add(cad);
            }
        }
            driver.close();
            return cadNumberList;
    }
}
