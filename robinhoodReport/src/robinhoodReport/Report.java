package robinhoodReport;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class Report {
    private static WebDriver driver;

    private static int rowNumber = 0;
    private static final int dateColumn = 0;
    private static final int brandNameColumn = 1;
    private static final int brandCodeColumn = 2;
    private static final int tradeTypeColumn = 3;
    private static final int countColumn = 4;
    private static final int tradePriceColumn = 5;
    private static final int totalPriceColumn = 6;

    private static File file;
    private static XSSFWorkbook wb;
    private static XSSFSheet sh;

    public static void main(String[] args) throws Exception {
        System.out.println("Report open");
        System.setProperty("webdriver.chrome.driver", "path/to/chromedriver");
        driver = new ChromeDriver();
        driver.get("http://localhost:3000/");

        file = new File("path/to/output.xlsx");
        wb = new XSSFWorkbook();
        sh = wb.createSheet();

        sh.createRow(rowNumber).createCell(dateColumn).setCellValue("Date");
        sh.getRow(rowNumber).createCell(brandNameColumn).setCellValue("Brand Name");
        sh.getRow(rowNumber).createCell(brandCodeColumn).setCellValue("Brand Code");
        sh.getRow(rowNumber).createCell(tradeTypeColumn).setCellValue("Buy / Sell");
        sh.getRow(rowNumber).createCell(countColumn).setCellValue("Count");
        sh.getRow(rowNumber).createCell(tradePriceColumn).setCellValue("Price");
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue("Total");

        List<WebElement> tradeRowsLocator = driver.findElements(By.xpath("//header/div"));
        for (WebElement tradeRowLocator : tradeRowsLocator) {
            WebElement leftSectionLocator = tradeRowLocator.findElement(By.xpath("div[1]"));
            WebElement rightSectionLocator = tradeRowLocator.findElement(By.xpath("div[2]"));
            WebElement rightTopSectionLocator = rightSectionLocator.findElement(By.tagName("h3"));
            String getTotalPriceText = rightTopSectionLocator.getText();

            Thread.sleep(2);
            if (!getTotalPriceText.equals("Placed") &&
                    !getTotalPriceText.equals("Canceled") &&
                    !getTotalPriceText.equals("$0.00") &&
                    !getTotalPriceText.equals("Voided") &&
                    !getTotalPriceText.equals("Failed") &&
                    !getTotalPriceText.equals("Rejected") &&
                    !getTotalPriceText.equals("Unable to fill")) {

                WebElement leftTopSectionLocator = leftSectionLocator.findElement(By.tagName("h3"));
                String leftTopSectionText = leftTopSectionLocator.getText();

                String leftTopSection = leftTopSectionLocator.getText();
                List<String> leftTopSectionSplits = new ArrayList<>(Arrays.asList(leftTopSection.split(" ")));

                String leftSection = leftSectionLocator.getText();
                List<String> leftSectionSplits = new ArrayList<>(Arrays.asList(leftSection.split("\n")));
                String getDate = leftSectionSplits.get(leftSectionSplits.size() - 1);

                String rightTopSection = rightTopSectionLocator.getText().replace("$", "").replace(",", "");
                double getTotalPrice = Double.parseDouble(rightTopSection);

                // option assignment
                if (leftTopSectionText.contains("Assignment")) {
                    String getBrandCode = leftTopSectionSplits.get(0).trim();
                    String getTradeType = leftTopSectionSplits.get(leftTopSectionSplits.size() - 1).trim();
                    String optionTradePrice = leftTopSectionSplits.get(1).trim().replace("$", "").replace(",", "");
                    double getOptionTradePrice = Double.parseDouble(optionTradePrice);
                    optionAssignment(getTotalPrice, getDate, getBrandCode, getOptionTradePrice);
                }
                // dividend
                else if (leftTopSectionText.contains("Dividend")) {
//                    leftTopSectionSplits.remove(leftTopSectionSplits.size() - 1);
//                    leftTopSectionSplits.remove(leftTopSectionSplits.size() - 1);
//                    String getBrandName = leftTopSectionSplits.stream().map(String::valueOf).collect(Collectors.joining(" ", "", ""));
//                    dividends(getDate, getBrandName, getTotalPrice);
                }
                // contract
                else if (leftTopSectionText.contains("Call") || leftTopSectionText.contains("Put")) {
                    WebElement rightBottomSectionLocator = rightSectionLocator.findElement(By.tagName("span"));
                    String rightBottomSection = rightBottomSectionLocator.getText();
                    List<String> rightBottomSectionSplits = new ArrayList<>(Arrays.asList(rightBottomSection.split(" ")));
                    String getBrandCode = leftTopSectionSplits.get(0).trim();
                    String getTradeType = leftTopSectionSplits.get(leftTopSectionSplits.size() - 1).trim();
                    double getCount = Double.parseDouble(rightBottomSectionSplits.get(0));
                    String tradePrice = rightBottomSectionSplits.get(rightBottomSectionSplits.size() - 1).replace("$", "").replace(",", "");
                    double getTradePrice = Double.parseDouble(tradePrice);
                    contracts(getTradeType, getTotalPrice, getDate, getBrandCode, getCount, getTradePrice);
                }
                // shares
                else if (leftTopSectionText.contains("Market") || leftTopSectionText.contains("Limit")) {
                    WebElement rightBottomSectionLocator = rightSectionLocator.findElement(By.tagName("span"));
                    String rightBottomSection = rightBottomSectionLocator.getText();
                    List<String> rightBottomSectionSplits = new ArrayList<>(Arrays.asList(rightBottomSection.split(" ")));
                    String getTradeType = leftTopSectionSplits.get(leftTopSectionSplits.size() - 1).trim();
                    leftTopSectionSplits.remove(leftTopSectionSplits.size() - 1);
                    leftTopSectionSplits.remove(leftTopSectionSplits.size() - 1);
                    String getBrandName = leftTopSectionSplits.stream().map(String::valueOf).collect(Collectors.joining(" ", "", ""));
                    String count = rightBottomSectionSplits.get(0).replace(",", "");
                    double getCount = Double.parseDouble(count);
                    String tradePrice = rightBottomSectionSplits.get(rightBottomSectionSplits.size() - 1).replace("$", "").replace(",", "");
                    double getTradePrice = Double.parseDouble(tradePrice);
                    shares(getTradeType, getTotalPrice, getDate, getBrandName, getCount, getTradePrice);
                }
//                // transfers
//                else if (leftTopSectionText.contains("Deposit") || leftTopSectionText.contains("Withdrawal")) {
//                    String getBrandCode = leftTopSectionSplits.get(0).trim();
//                    transfers(getDate, getBrandCode, getTotalPrice);
//                }
//                // rewards
//                else if (leftTopSectionText.contains("Referral")) {
//                    rewards(getDate, getTotalPrice);
//                }

                try {
                    FileOutputStream fos = new FileOutputStream(file);
                    wb.write(fos);
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        }

        driver.close();
        driver.quit();
        System.out.println("Report close");
    }

    private static void optionAssignment(double getTotalPrice, String getDate, String  getBrandCode, double getOptionTradePrice) {
        String tradeType = (getTotalPrice > 0.0) ? "Sell" : "Buy";
        sh.createRow(rowNumber).createCell(dateColumn).setCellValue(getDate);
        sh.getRow(rowNumber).createCell(brandCodeColumn).setCellValue(getBrandCode);
        sh.getRow(rowNumber).createCell(tradeTypeColumn).setCellValue(tradeType);
        sh.getRow(rowNumber).createCell(countColumn).setCellValue(Math.abs(Math.ceil(getTotalPrice / getOptionTradePrice)));
        sh.getRow(rowNumber).createCell(tradePriceColumn).setCellValue(getOptionTradePrice);
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue(Math.ceil(getTotalPrice));
    }

    private static void contracts(String getTradeType, double getTotalPrice, String getDate, String getBrandCode, double getCount, double getTradePrice) {
        double totalPrice = getTradeType.equals("Buy") ? getTotalPrice * -1 : getTotalPrice;
        sh.createRow(rowNumber).createCell(dateColumn).setCellValue(getDate);
        sh.getRow(rowNumber).createCell(brandCodeColumn).setCellValue(getBrandCode);
        sh.getRow(rowNumber).createCell(tradeTypeColumn).setCellValue(getTradeType);
        sh.getRow(rowNumber).createCell(countColumn).setCellValue(getCount * 100);
        sh.getRow(rowNumber).createCell(tradePriceColumn).setCellValue(getTradePrice);
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue(totalPrice);
    }

    private static void shares(String getTradeType, double getTotalPrice, String getDate, String getBrandName, double getCount, double getTradePrice) {
        double totalPrice = getTradeType.equals("Buy") ? getTotalPrice * -1 : getTotalPrice;
        sh.createRow(rowNumber).createCell(dateColumn).setCellValue(getDate);
        sh.getRow(rowNumber).createCell(brandNameColumn).setCellValue(getBrandName);
        sh.getRow(rowNumber).createCell(tradeTypeColumn).setCellValue(getTradeType);
        sh.getRow(rowNumber).createCell(countColumn).setCellValue(getCount);
        sh.getRow(rowNumber).createCell(tradePriceColumn).setCellValue(getTradePrice);
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue(totalPrice);
    }

    private static void dividends(String getDate, String getBrandName, double getTotalPrice) {
        sh.createRow(rowNumber).createCell(dateColumn).setCellValue(getDate);
        sh.getRow(rowNumber).createCell(brandNameColumn).setCellValue(getBrandName);
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue(getTotalPrice);
    }

    private static void transfers(String getDate, String getBrandCode, double getTotalPrice) {
        sh.createRow(rowNumber).createCell(dateColumn).setCellValue(getDate);
        sh.getRow(rowNumber).createCell(brandCodeColumn).setCellValue(getBrandCode);
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue(getTotalPrice);
    }

    private static void rewards(String getDate, double getTotalPrice) {
        sh.createRow(rowNumber).createCell(dateColumn).setCellValue(getDate);
        sh.getRow(rowNumber++).createCell(totalPriceColumn).setCellValue(getTotalPrice);
    }
}
