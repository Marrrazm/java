package com.beginsecure;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

public class ExpenseTracker {

    private static final Logger logger = LogManager.getLogger(ExpenseTracker.class);

    private final Map<String, Map<LocalDate, List<Double>>> expensesByCategory = new HashMap<>();
    private final List<String> categories = new ArrayList<>(Arrays.asList(
            "Food", "Rent", "Transport", "Clothing", "Internet", "Beautiful",
            "Marketplaces", "Nalogi", "Health", "Gifts"
    ));

    public ExpenseTracker() {
        for (String category : categories) {
            expensesByCategory.put(category, new HashMap<>());
        }
    }

    public void addCategory(String newCategory) {
        if (categories.contains(newCategory)) {
            logger.warn("Category already exists: " + newCategory);
            throw new IllegalArgumentException("Category already exists: " + newCategory);
        }
        categories.add(newCategory);
        expensesByCategory.put(newCategory, new HashMap<>());
        logger.info("Added new category: " + newCategory);
    }

    public void addExpense(String category, double amount, LocalDate date) {
        if (!categories.contains(category)) {
            logger.warn("Invalid category: " + category);
            throw new IllegalArgumentException("Invalid category: " + category);
        }
        expensesByCategory.computeIfAbsent(category, k -> new HashMap<>())
                .computeIfAbsent(date, k -> new ArrayList<>()).add(amount);
        logger.info("Added expense: {} to category {} on {}", amount, category, date);
    }

    public void showStatistics(Optional<String> category) {
        if (category.isPresent() && !categories.contains(category.get())) {
            logger.warn("Invalid category: " + category.get());
            throw new IllegalArgumentException("Invalid category: " + category.get());
        }

        if (category.isEmpty()) {
            double totalSpent = expensesByCategory.values().stream()
                    .flatMap(map -> map.values().stream().flatMap(List::stream))
                    .mapToDouble(Double::doubleValue).sum();
            System.out.println("Total spent: " + totalSpent);

            for (String cat : categories) {
                double totalByCategory = expensesByCategory.get(cat).values().stream()
                        .flatMap(List::stream)
                        .mapToDouble(Double::doubleValue).sum();
                double percentage = totalSpent > 0 ? (totalByCategory / totalSpent) * 100 : 0;
                System.out.printf("%s: %.2f (%.2f%%)%n", cat, totalByCategory, percentage);
            }
        } else {
            String cat = category.get();
            double totalByCategory = expensesByCategory.get(cat).values().stream()
                    .flatMap(List::stream)
                    .mapToDouble(Double::doubleValue).sum();
            System.out.printf("%s: %.2f%n", cat, totalByCategory);
        }
    }

    public void exportToExcel(Optional<String> category, String fileName) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Expenses");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Date");
            header.createCell(1).setCellValue("Category");
            header.createCell(2).setCellValue("Amount");

            int rowNum = 1;
            for (String cat : categories) {
                if (category.isPresent() && !category.get().equals(cat)) {
                    continue;
                }

                for (Map.Entry<LocalDate, List<Double>> entry : expensesByCategory.get(cat).entrySet()) {
                    for (Double amount : entry.getValue()) {
                        Row row = sheet.createRow(rowNum++);
                        row.createCell(0).setCellValue(entry.getKey().toString());
                        row.createCell(1).setCellValue(cat);
                        row.createCell(2).setCellValue(amount);
                    }
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
                workbook.write(outputStream);
                logger.info("Exported data to Excel file: " + fileName);
            }
        } catch (IOException e) {
            logger.error("Error writing Excel file", e);
        }
    }

    public static void main(String[] args) {
        ExpenseTracker tracker = new ExpenseTracker();
        Scanner scanner = new Scanner(System.in);

        while (true) {
            System.out.println("1. Add Expense\n2. Add Category\n3. Show Statistics\n4. Export to Excel\n5. Exit");
            int choice = scanner.nextInt();
            scanner.nextLine();

            switch (choice) {
                case 1 -> {
                    System.out.println("Enter category:");
                    String category = scanner.nextLine();
                    System.out.println("Enter amount:");
                    double amount = scanner.nextDouble();
                    System.out.println("Enter date (yyyy-mm-dd):");
                    LocalDate date = LocalDate.parse(scanner.next());
                    tracker.addExpense(category, amount, date);
                }
                case 2 -> {
                    System.out.println("Enter new category name:");
                    String newCategory = scanner.nextLine();
                    tracker.addCategory(newCategory);
                }
                case 3 -> {
                    System.out.println("Enter category (or press Enter for all):");
                    String category = scanner.nextLine();
                    tracker.showStatistics(category.isEmpty() ? Optional.empty() : Optional.of(category));
                }
                case 4 -> {
                    System.out.println("Enter category (or press Enter for all):");
                    String category = scanner.nextLine();
                    System.out.println("Enter file name:");
                    String fileName = scanner.nextLine();
                    tracker.exportToExcel(category.isEmpty() ? Optional.empty() : Optional.of(category), fileName);
                }
                case 5 -> {
                    logger.info("Exiting application.");
                    return;
                }
                default -> System.out.println("Invalid choice.");
            }
        }
    }
}
```
