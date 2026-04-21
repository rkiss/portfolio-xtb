package name.abuchen.portfolio.datatransfer.xtb;

import java.time.LocalDateTime;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

/**
 * Utilities around the Apache POI XSSF component
 */
public class PoiXSSFUtils
{

    private PoiXSSFUtils()
    {
        // avoid instantiation
    }

    private static String getCellRef(Row row, String colIndex)
    {
        return String.format("'%s.%s%s'", row.getSheet().getSheetName(), colIndex, row.getRowNum());
    }

    public static String getCellValueAsString(Row row, String colIndex, String defaultValue, List<Exception> errors)
    {
        try
        {
            Cell cell = row.getCell(CellReference.convertColStringToIndex(colIndex));
            if (cell == null)
                throw new IllegalStateException("Null value");
            switch (cell.getCellType())
            {
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return cell.getCellFormula();
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case STRING:
                    return cell.getStringCellValue();
                default:
                    return defaultValue;
            }
        }
        catch (Exception e)
        {
            if (errors != null)
            {
                errors.add(new IllegalStateException(
                                "Error on cell " + getCellRef(row, colIndex) + ": " + e.getMessage()));
            }
            return defaultValue;
        }
    }

    public static long getCellValueAsLong(Row row, String colIndex, long defaultValue, List<Exception> errors)
    {
        try
        {
            Cell cell = row.getCell(CellReference.convertColStringToIndex(colIndex));
            if (cell == null)
                throw new IllegalStateException("Null value");
            switch (cell.getCellType())
            {
                case BOOLEAN:
                    return cell.getBooleanCellValue() ? 1 : 0;
                case FORMULA, NUMERIC:
                    return (long) cell.getNumericCellValue();
                case STRING:
                    return Double.valueOf(cell.getStringCellValue()).longValue();
                default:
                    return defaultValue;
            }
        }
        catch (Exception e)
        {
            if (errors != null)
            {
                errors.add(new IllegalStateException(
                                "Error on cell " + getCellRef(row, colIndex) + ": " + e.getMessage()));
            }
            return defaultValue;
        }
    }

    public static double getCellValueAsDouble(Row row, String colIndex, double defaultValue, List<Exception> errors)
    {
        try
        {
            Cell cell = row.getCell(CellReference.convertColStringToIndex(colIndex));
            if (cell == null)
                throw new IllegalStateException("Null value");
            switch (cell.getCellType())
            {
                case BOOLEAN:
                    return cell.getBooleanCellValue() ? 1 : 0;
                case FORMULA, NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    return Double.valueOf(cell.getStringCellValue());
                default:
                    return defaultValue;
            }
        }
        catch (Exception e)
        {
            if (errors != null)
            {
                errors.add(new IllegalStateException(
                                "Error on cell " + getCellRef(row, colIndex) + ": " + e.getMessage()));
            }
            return defaultValue;
        }
    }

    public static LocalDateTime getCellValueAsLocalDateTime(Row row, String colIndex, LocalDateTime defaultValue,
                    List<Exception> errors)
    {
        try
        {
            Cell cell = row.getCell(CellReference.convertColStringToIndex(colIndex));
            if (cell == null)
                throw new IllegalStateException("Null value");
            switch (cell.getCellType())
            {
                case NUMERIC:
                    LocalDateTime result = cell.getLocalDateTimeCellValue();
                    if (result == null)
                        throw new IllegalStateException("Cannot convert empty value to date");
                    else
                        return result;
                default:
                    return defaultValue;

            }
        }
        catch (Exception e)
        {
            if (errors != null)
            {
                errors.add(new IllegalStateException(
                                "Error on cell " + getCellRef(row, colIndex) + ": " + e.getMessage()));
            }
            return defaultValue;
        }
    }

}
