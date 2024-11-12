package weeklyfilecompare.newapproach;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class WeeklyFileComparator {

	private static final String SCRUB_RULES_FILE = "resources/Scrub rules.xlsx";
	// Columns to use for comparison
	private static final String[] COMPARE_COLUMNS = { "Proc_code", "Modifiers", "CMSAdd", "CMSTerm", "Service",
			"Service desc", "RateType", "Pricing Method", "Rate Eff", "Rate Term", "MAxFee" };

	// Columns to write into the output sheet
	private static final String[] OUTPUT_COLUMNS = { "Proc_code", "Modifiers", "CMSAdd", "CMSTerm", "Rate Eff",
			"Rate Term", "MAxFee" };

	public static void main(String[] args) throws IOException {
		String lastWeekFile = "C:\\Users\\rajas\\Desktop\\Excelcompare\\Lastweekfile.xlsx";
		String currentWeekFile = "C:\\Users\\rajas\\Desktop\\Excelcompare\\Thisweekfile.xlsx";
		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
		String outputFile = "C:\\Users\\rajas\\Desktop\\Excelcompare\\OutputReport" + timestamp + ".xlsx";
		// Step 1: Load data from both files
		List<Map<String, String>> lastWeekData = loadData(lastWeekFile);
		List<Map<String, String>> currentWeekData = loadData(currentWeekFile);

		// Step 2: Find exact matches using retainAll
		List<Map<String, String>> exactMatch = new ArrayList<>(currentWeekData);
		exactMatch.retainAll(lastWeekData);

		// Step 3: Get differences by removing exact matches from current week data
		List<Map<String, String>> diffMatch = new ArrayList<>(currentWeekData);
		diffMatch.removeAll(exactMatch);

		List<Map<String, Object>> scrubRules = loadScrubRules(SCRUB_RULES_FILE);

		// Step 4: Write differences to the output report
		try (Workbook outputWorkbook = new XSSFWorkbook()) {
			Sheet outputSheet = outputWorkbook.createSheet("Mismatch Records");
			writeDifferencesToSheet(outputSheet, diffMatch, lastWeekData, currentWeekData, scrubRules);

			try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
				outputWorkbook.write(fileOut);
				System.out.println("Comparison completed. Report generated at: " + outputFile);
			}
		}
	}

	// Step 1: Load data from the Excel file into a list of maps
	private static List<Map<String, String>> loadData(String filePath) throws IOException {
		List<Map<String, String>> data = new ArrayList<>();
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath))) {
			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);

			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				Map<String, String> rowData = new HashMap<>();

				for (String column : COMPARE_COLUMNS) {
					int colIndex = getColumnIndex(headerRow, column);
					rowData.put(column, getCellValue(row.getCell(colIndex)));
				}
				data.add(rowData);
			}
		}
		return data;
	}

	// Step 4: Write mismatched records to the output sheet with LastWeek and
	private static void writeDifferencesToSheet(Sheet outputSheet, List<Map<String, String>> diffMatch,
			List<Map<String, String>> lastWeekData, List<Map<String, String>> currentWeekData,
			List<Map<String, Object>> scrubRules) {
		Row headerRow = outputSheet.createRow(0);
		int colIdx = 0;

		// Write the initial columns based on OUTPUT_COLUMNS
		for (String column : OUTPUT_COLUMNS) {
			headerRow.createCell(colIdx++).setCellValue(column);
		}

		// Define specific column indexes for readability
		int differencesColIdx = colIdx++;
		int scrubColIdx = colIdx++;
		int ruleDescriptionColIdx = colIdx++;

		// Write header labels for specific columns
		headerRow.createCell(differencesColIdx).setCellValue("Differences");
		headerRow.createCell(scrubColIdx).setCellValue("Scrub");
		headerRow.createCell(ruleDescriptionColIdx).setCellValue("Rule Description");

		Map<String, Integer> dynamicHeaderMap = new HashMap<>(); // To track dynamic headers for mismatched columns
		int additionalColumnsStartIndex = colIdx;
		int outputRowNum = 1;

		// Step 1: Write New Codes and Other Differences from Thisweekfile
		for (Map<String, String> currentRow : diffMatch) {
			String procCode = currentRow.get("Proc_code");
			String modifiers = currentRow.get("Modifiers");

			// Check if the Proc_code + Modifiers combination exists in lastWeekData
			Map<String, String> lastWeekRow = lastWeekData.stream()
					.filter(row -> row.get("Proc_code").equals(procCode) && row.get("Modifiers").equals(modifiers))
					.findFirst().orElse(null);

			Row outputRow = outputSheet.createRow(outputRowNum++);
			int diffColIdx = 0;

			// Write initial OUTPUT_COLUMNS values from current week data
			for (String column : OUTPUT_COLUMNS) {
				outputRow.createCell(diffColIdx++).setCellValue(currentRow.get(column));
			}

			if (lastWeekRow == null) {
				// If lastWeekRow is null, this is a new code
				outputRow.createCell(diffColIdx).setCellValue("New code");
				outputRow.createCell(scrubColIdx).setCellValue(""); // Scrub column
				outputRow.createCell(ruleDescriptionColIdx).setCellValue(""); // Rule Description column
			} else {
				// If lastWeekRow exists, perform a column-by-column comparison
				StringBuilder differences = new StringBuilder();
				Map<String, String> mismatchedValues = new LinkedHashMap<>();

				for (String column : COMPARE_COLUMNS) {
					String lastValue = lastWeekRow.get(column);
					String currentValue = currentRow.get(column);

					// If there’s a mismatch, store it for later writing
					if (!Objects.equals(lastValue, currentValue)) {
						differences.append(column).append(", ");
						mismatchedValues.put("LastWeek." + column, lastValue);
						mismatchedValues.put("ThisWeek." + column, currentValue);
					}
				}

				if (differences.length() > 0) {
					outputRow.createCell(diffColIdx++).setCellValue(differences.substring(0, differences.length() - 2));

					outputRow.createCell(scrubColIdx).setCellValue(""); // Scrub column initially blank
					outputRow.createCell(ruleDescriptionColIdx).setCellValue(""); // Rule Description column initially
					// blank

					int additionalColIdx = colIdx;
					// Write mismatched values under appropriate LastWeek and ThisWeek headers
					for (Map.Entry<String, String> entry : mismatchedValues.entrySet()) {
						String header = entry.getKey();
						int headerIndex;

						// Check if header already exists; if not, create it and track the index
						if (!dynamicHeaderMap.containsKey(header)) {
							headerIndex = headerRow.getLastCellNum();
							headerRow.createCell(headerIndex).setCellValue(header);
							dynamicHeaderMap.put(header, headerIndex);
						} else {
							headerIndex = dynamicHeaderMap.get(header);
						}

						outputRow.createCell(headerIndex).setCellValue(entry.getValue());
					}
				}
			}

			for (Map<String, Object> rule : scrubRules) {
				Map<String, String> conditions = (Map<String, String>) rule.get("conditions");
				boolean matches = conditions.entrySet().stream()
						.allMatch(entry -> Objects.equals(currentRow.get(entry.getKey()), entry.getValue()));

				if (matches) {
					outputRow.getCell(scrubColIdx).setCellValue("Yes"); // Set Scrub column to "Yes"
					outputRow.getCell(ruleDescriptionColIdx).setCellValue((String) rule.get("Rule Description"));

					break; // Stop after the first matching rule
				}
			}
		}

		// Step 3: Write Termed Codes from Lastweekfile that do not exist in
		// Thisweekfile
		for (Map<String, String> lastRow : lastWeekData) {
			String procCode = lastRow.get("Proc_code");
			String modifiers = lastRow.get("Modifiers");

			// Check if the Proc_code + Modifiers combination exists in currentWeekData
			boolean existsInThisWeek = currentWeekData.stream()
					.anyMatch(row -> row.get("Proc_code").equals(procCode) && row.get("Modifiers").equals(modifiers));

			if (!existsInThisWeek) {
				// If it doesn't exist in this week, it's a termed code
				Row outputRow = outputSheet.createRow(outputRowNum++);
				int diffColIdx = 0;

				// Write initial OUTPUT_COLUMNS values from last week data for termed codes
				for (String column : OUTPUT_COLUMNS) {
					outputRow.createCell(diffColIdx++).setCellValue(lastRow.get(column));
				}

				// Mark the Differences column as "Termed code"
				outputRow.createCell(diffColIdx).setCellValue("Termed code");

				// Initialize Scrub and Rule Description columns as empty
				outputRow.createCell(scrubColIdx).setCellValue("");
				outputRow.createCell(ruleDescriptionColIdx).setCellValue("");

				for (Map<String, Object> rule : scrubRules) {
					@SuppressWarnings("unchecked")
					Map<String, String> conditions = (Map<String, String>) rule.get("conditions");

					// Check if all conditions in the rule match the lastRow
					boolean matches = conditions.entrySet().stream()
							.allMatch(entry -> Objects.equals(lastRow.get(entry.getKey()), entry.getValue()));

					if (matches) {
						// Set Scrub column to "Yes" and populate the Rule Description column
						outputRow.getCell(scrubColIdx).setCellValue("Yes");
						outputRow.getCell(ruleDescriptionColIdx).setCellValue((String) rule.get("Rule Description"));
						break; // Stop after the first matching rule
					}
				}

			}

		}
	}

	private static List<Map<String, Object>> loadScrubRules(String filePath) throws IOException {
		List<Map<String, Object>> scrubRules = new ArrayList<>();
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath))) {
			Sheet sheet = workbook.getSheetAt(0);
			Row headerRow = sheet.getRow(0);

			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				Map<String, String> conditions = new HashMap<>();

				// Read the Rule Description from column B
				Cell ruleDescriptionCell = row.getCell(1); // Assuming column B is index 1
				String ruleDescription = getCellValue(ruleDescriptionCell);

				// Read the conditions from columns C to F
				for (int j = 2; j <= 5; j++) { // Columns C to F
					Cell cell = row.getCell(j);
					if (cell != null && cell.getCellType() != CellType.BLANK) {
						String header = getCellValue(headerRow.getCell(j));
						String value = getCellValue(cell);
						conditions.put(header, value);
					}
				}

				if (!conditions.isEmpty()) {
					Map<String, Object> rule = new HashMap<>();
					rule.put("conditions", conditions); // Store the conditions map
					rule.put("Rule Description", ruleDescription); // Store Rule Description separately
					scrubRules.add(rule);
				}
			}
		}
		return scrubRules;
	}

	// Helper to get column index by header name
	private static int getColumnIndex(Row headerRow, String columnName) {
		for (int i = 0; i < headerRow.getLastCellNum(); i++) {
			if (headerRow.getCell(i).getStringCellValue().equalsIgnoreCase(columnName)) {
				return i;
			}
		}
		return -1;
	}

	// Helper to get cell value as string
	private static String getCellValue(Cell cell) {
		if (cell == null)
			return "";
		return switch (cell.getCellType()) {
		case STRING -> cell.getStringCellValue().trim();
		case NUMERIC -> String.valueOf(cell.getNumericCellValue()).trim();
		case BOOLEAN -> String.valueOf(cell.getBooleanCellValue()).trim();
		default -> "";
		};
	}
}