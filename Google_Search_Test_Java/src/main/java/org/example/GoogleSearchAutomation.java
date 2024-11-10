package org.example;

import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Collections;
import java.util.Comparator;
import java.util.stream.Collectors;
import java.time.DayOfWeek;
import java.time.LocalDate;

public class GoogleSearchAutomation {

    public static void main(String[] args) {
        // Set the path to the ChromeDriver executable
        System.setProperty("webdriver.chrome.driver", "src/main/resources/driver/chromedriver.exe");

        // Initialize the WebDriver
        WebDriver driver = new ChromeDriver();

        // Initialize WebDriverWait
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));

        try {
            // Open Google
            driver.get("https://www.google.com");

            // Get today's day name
            DayOfWeek currentDay = LocalDate.now().getDayOfWeek();
            String currentDayName = currentDay.name().substring(0, 1) + currentDay.name().substring(1).toLowerCase();

            // Get the keywords for today
            List<String> keywords = ExcelHandler.getKeywordsForDay(currentDayName);

            // Iterate through the keywords and perform the Google search
            for (String keyword : keywords) {
                // Locate the search box using its name attribute
                WebElement searchBox = driver.findElement(By.name("q"));

                // Clear the search box if any previous value exists
                searchBox.clear();

                // Type the keyword into the search box
                searchBox.sendKeys(keyword);
                Thread.sleep(2000); // Wait for suggestions to load

                // Wait for suggestions to appear
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div#Alh6id ul li div.wM6W7d span")));

                // Locate suggestion elements
                List<WebElement> suggestions = driver.findElements(By.cssSelector("div#Alh6id ul li div.wM6W7d span"));

                String longestSuggestion = "";
                String smallestSuggestion = "";

                if (!suggestions.isEmpty()) {
                    // Extract text from suggestions
                    List<String> suggestionTexts = suggestions.stream()
                            .map(WebElement::getText)
                            .filter(text -> text != null && !text.isEmpty())  // Filter out null or empty strings
                            .collect(Collectors.toList());

                    if (!suggestionTexts.isEmpty()) {
                        // Find longest and smallest suggestions
                        longestSuggestion = Collections.max(suggestionTexts, Comparator.comparingInt(String::length));
                        smallestSuggestion = Collections.min(suggestionTexts, Comparator.comparingInt(String::length));
                    } else {
                        longestSuggestion = "No valid suggestions";
                        smallestSuggestion = "No valid suggestions";
                    }
                } else {
                    longestSuggestion = "No suggestions found";
                    smallestSuggestion = "No suggestions found";
                }

                // Print the result (optional)
                System.out.println("Keyword: " + keyword);
                System.out.println("Longest Suggestion: " + longestSuggestion);
                System.out.println("Smallest Suggestion: " + smallestSuggestion);

                // Write the result to Excel with the dynamic sheet name
                ExcelHandler.writeResultsToExcel(currentDayName, keyword, longestSuggestion, smallestSuggestion);

                // Optionally wait before the next iteration
                Thread.sleep(2000); // Wait for the next keyword
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Close the browser after performing all searches
            driver.quit();
        }
    }
}
