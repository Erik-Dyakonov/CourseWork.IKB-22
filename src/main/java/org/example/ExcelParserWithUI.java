package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.general.DefaultPieDataset;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.*;
import java.util.List;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ExcelParserWithUI extends JFrame {

    @Serial
    private static final long serialVersionUID = 1L;

    //Создание GUI
    public ExcelParserWithUI() {
        super("Анализатор Excel");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new FlowLayout());

        // Создание кнопки "Выбрать и анализировать Excel"
        JButton analyzeButton = new JButton("Выбрать и анализировать Excel");
        analyzeButton.addActionListener(e -> performAnalysis());

        // Добавление кнопки на главное окно
        add(analyzeButton);
    }

    // Метод для выполнения анализа
    private void performAnalysis() {
        try {
            // Создание диалога выбора файла
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Выберите Excel-файл");
            fileChooser.setFileFilter(new FileNameExtensionFilter("Файлы Excel", "xlsx"));

            // Отображение диалога и обработка выбора пользователя
            int userSelection = fileChooser.showOpenDialog(this);
            if (userSelection != JFileChooser.APPROVE_OPTION) {
                return;
            }

            // Получение выбранного файла
            File selectedFile = fileChooser.getSelectedFile();

            // Проверка расширения файла
            if (!selectedFile.getName().toLowerCase().endsWith(".xlsx")) {
                showErrorDialog("Выбранный файл не является файлом Excel (.xlsx).");
                return;
            }

            // Загрузка исходного файла Excel
            Workbook workbook = new XSSFWorkbook(new FileInputStream(selectedFile));
            Sheet sheet = workbook.getSheetAt(0);

            // Инициализация структур для хранения статистики и категорий студентов
            Map<String, Integer> gradeCount = new HashMap<>();
            Map<String, List<String>> excellentStudents = new HashMap<>();
            Map<String, List<String>> goodStudents = new HashMap<>();
            Map<String, List<String>> satisfactoryStudents = new HashMap<>();
            Map<String, List<String>> notAllowedStudents = new HashMap<>();

            int totalStudents = 0;
            double totalScore = 0;
            double maxScore = 5;

            // Обработка данных из файла
            for (Row row : sheet) {
                Cell nameCell = row.getCell(0);
                Cell scoreCell = row.getCell(1);

                if (nameCell != null && scoreCell != null) {
                    String name = nameCell.getStringCellValue();
                    double score = scoreCell.getNumericCellValue();

                    // Обновление статистики
                    gradeCount.put(name, (int) score);

                    // Разделение студентов по категориям
                    String scoreCategory = getScoreCategory(score);
                    addToCategoryMap(scoreCategory, name, excellentStudents, goodStudents, satisfactoryStudents, notAllowedStudents);

                    totalStudents++;
                    totalScore += score;
                }
            }

            // Рассчет статистики по оценкам
            int excellentCount = countStudentsByScore(gradeCount.values(), maxScore);
            int goodCount = countStudentsByScore(gradeCount.values(), 4.0);
            int satisfactoryCount = countStudentsByScore(gradeCount.values(), 3.0);
            int notAllowedCount = countStudentsByScore(gradeCount.values(), 2.0);
            double averageScore = totalScore / totalStudents;

            // Создание нового Excel-файла с результатами
            Workbook resultWorkbook = new XSSFWorkbook();
            Sheet resultSheet = resultWorkbook.createSheet("Результаты");

            // Запись заголовков статистики в новый файл
            int rowNumber = 0;
            rowNumber = writeStatisticRow(resultSheet, rowNumber, "Отличники", excellentCount);
            rowNumber = writeStatisticRow(resultSheet, rowNumber, "Хорошисты", goodCount);
            rowNumber = writeStatisticRow(resultSheet, rowNumber, "Троечники", satisfactoryCount);
            rowNumber = writeStatisticRow(resultSheet, rowNumber, "Не допущены", notAllowedCount);
            rowNumber = writeStatisticRow(resultSheet, rowNumber, "Средний балл", averageScore);

            // Добавление информации о студентах в категориях
            rowNumber = writeCategoryResults(resultSheet, "Отличники", excellentStudents, rowNumber);
            rowNumber = writeCategoryResults(resultSheet, "Хорошисты", goodStudents, rowNumber);
            rowNumber = writeCategoryResults(resultSheet, "Троечники", satisfactoryStudents, rowNumber);
            rowNumber = writeCategoryResults(resultSheet, "Не допущены", notAllowedStudents, rowNumber);

            // Сохранение нового файла
            saveWorkbook(resultWorkbook);

            // Создание круговой диаграммы на основе статистики
            JFreeChart pieChart = createPieChart(notAllowedCount, satisfactoryCount, goodCount, excellentCount);

            // Отображение диаграммы в отдельном окне
            displayChart(pieChart);

            // Закрытие исходного файла
            workbook.close();

            // Оповещение об успешном завершении
            showInfoDialog();
        } catch (IOException | IllegalStateException e) {
            // Обработка исключений
            handleException(e);
        }
    }

    // Получение категории оценки по числовому значению
    private String getScoreCategory(double score) {
        if (score == 5.0) return "Отличники";
        else if (score == 4.0) return "Хорошисты";
        else if (score == 3.0) return "Троечники";
        else return "Не допущены";
    }

    // Добавление студента в указанную категорию
    @SafeVarargs
    private void addToCategoryMap(String category, String studentName, Map<String, List<String>>... studentsMaps) {
        for (Map<String, List<String>> studentsMap : studentsMaps) {
            addToMap(studentsMap, category, studentName);
        }
    }

    // Добавление значения в карту по ключу, если ключ отсутствует
    private void addToMap(Map<String, List<String>> map, String key, String value) {
        map.computeIfAbsent(key, k -> new ArrayList<>()).add(value);
    }

    // Подсчет студентов с определенной оценкой
    private int countStudentsByScore(Collection<Integer> scores, double targetScore) {
        return (int) scores.stream().filter(score -> score == targetScore).count();
    }

    // Запись результатов категории студентов в файл
    private int writeCategoryResults(Sheet sheet, String category, Map<String, List<String>> studentsMap, int rowNumber) {
        // Запись заголовка категории
        Row headerRow = sheet.createRow(rowNumber++);
        headerRow.createCell(0).setCellValue(category);

        // Запись ФИО в соответствующую категорию
        List<String> students = studentsMap.getOrDefault(category, Collections.emptyList());
        for (String student : students) {
            Row dataRow = sheet.createRow(rowNumber++);
            dataRow.createCell(0).setCellValue(student);
        }

        return rowNumber;
    }

    // Запись строки статистики в файл
    private int writeStatisticRow(Sheet sheet, int rowNumber, String label, double value) {
        Row row = sheet.createRow(rowNumber++);
        row.createCell(0).setCellValue(label);
        row.createCell(1).setCellValue(value);
        return rowNumber;
    }

    // Сохранение Excel-файла
    private void saveWorkbook(Workbook workbook) throws IOException {
        try (FileOutputStream fileOutputStream = new FileOutputStream("результаты.xlsx")) {
            workbook.write(fileOutputStream);
        }
    }

    // Создание круговой диаграммы на основе статистики
    private JFreeChart createPieChart(int notAllowedCount, int satisfactoryCount, int goodCount, int excellentCount) {
        // Создание набора данных для круговой диаграммы
        DefaultPieDataset<String> dataset = new DefaultPieDataset<>();
        dataset.setValue("Троечники", satisfactoryCount);
        dataset.setValue("Хорошисты", goodCount);
        dataset.setValue("Отличники", excellentCount);
        dataset.setValue("Не допущенные", notAllowedCount);

        // Создание круговой диаграммы с использованием JFreeChart
        JFreeChart chart = ChartFactory.createPieChart("Распределение оценок", dataset, true, true, false);
        // Получение объекта Plot (графического объекта) из диаграммы
        PiePlot plot = (PiePlot) chart.getPlot();
        // Установка простых (текстовых) меток на диаграмме
        plot.setSimpleLabels(true);

        return chart;
    }

    // Отображение диаграммы в окне
    private void displayChart(JFreeChart chart) {
        // Создание нового окна JFrame для отображения диаграммы
        JFrame chartFrame = new JFrame("Круговая диаграмма");
        // Установка менеджера компоновки BorderLayout для окна
        chartFrame.setLayout(new BorderLayout());
        // Создание панели с диаграммой
        ChartPanel chartPanel = new ChartPanel(chart);
        // Добавление панели на центральное место окна
        chartFrame.add(chartPanel, BorderLayout.CENTER);
        chartFrame.setSize(500, 500);
        // Установка положения окна по центру экрана
        chartFrame.setLocationRelativeTo(null);
        // Установка видимости окна (отображение окна)
        chartFrame.setVisible(true);
    }

    // Отображение диалогового окна с ошибкой
    private void showErrorDialog(String message) {
        JOptionPane.showMessageDialog(this, message, "Ошибка", JOptionPane.ERROR_MESSAGE);
    }

    // Отображение диалогового окна с информацией
    private void showInfoDialog() {
        JOptionPane.showMessageDialog(this, "Анализ завершен.", "Успех", JOptionPane.INFORMATION_MESSAGE);
    }

    // Обработка исключений и вывод в лог
    private void handleException(Exception e) {
        Logger.getLogger(ExcelParserWithUI.class.getName()).log(Level.SEVERE, "Произошла ошибка при анализе Excel.", e);
        showErrorDialog("Произошла ошибка: " + e.getMessage());
    }

    // Точка входа в программу
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            ExcelParserWithUI frame = new ExcelParserWithUI();
            frame.setSize(300, 100);
            frame.setLocationRelativeTo(null);
            frame.setVisible(true);
        });
    }
}
