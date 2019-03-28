package excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;


/**
 *
 * excl读取数据工具
 * @create 2018-05-02 15:28
 */
public class ExcelUtil
{

    private int totalRows = 0;
    
    private int totalCells = 0;
    
    private String errorInfo;
    
    public ExcelUtil()
    {
        
    }
    

    
    public int getTotalRows()
    {
        return totalRows;
    }
    
    public int getTotalCells()
    {
        return totalCells;
        
    }
    
    public String getErrorInfo()
    {
        return errorInfo;
    }
    
    public boolean validateExcel(String filePath)
    {
        
        if (filePath == null
                || !(isExcel2003(filePath) || isExcel2007(filePath)))
        {
            errorInfo = "";
            return false;
        }
        
        File file = new File(filePath);
        if (!file.exists())
        {
            errorInfo = "";
            return false;
        }
        
        return true;
        
    }
    
    public List<List<String>> read(String filePath, int validColumn)
    {
        
        List<List<String>> dataLst = new ArrayList<List<String>>();
        
        InputStream is = null;
        
        try
        {
            
            if (!validateExcel(filePath))
            {
                throw new Exception();
            }
            
            boolean isExcel2003 = true;
            if (isExcel2007(filePath))
            {
                isExcel2003 = false;
            }
            File file = new File(filePath);
            is = new FileInputStream(file);
            dataLst = read(is, isExcel2003, validColumn);
            is.close();
            
        }
        catch (Exception ex)
        {

        }
        finally
        {
            if (is != null)
            {
                try
                {
                    is.close();
                }
                catch (IOException e)
                {
                    is = null;
                }
            }
            
        }
        return dataLst;
    }
    
    /**
     * 读取excl
     * @param file	
     * @param validColumn 文件有效列数和传入文件的列数是否相等
     * @return
     * @throws
     */
    public List<List<String>> read(File file, int validColumn)
    {
    	
        if (file == null)
        {
            return null;
        }
        String filePath = file.getAbsolutePath();

        List<List<String>> dataLst = new ArrayList<List<String>>();
        
        InputStream is = null;
        
        try
        {
            if (!validateExcel(filePath))
            {
            }
            boolean isExcel2003 = true;
            if (isExcel2007(filePath))
            {
                isExcel2003 = false;
            }
            is = new FileInputStream(file);
            dataLst = read(is, isExcel2003, validColumn);
            is.close();
            
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
        }
        finally
        {
            if (is != null)
            {
                try
                {
                    is.close();
                }
                catch (IOException e)
                {
                    is = null;
                }
            }
            
        }
        return dataLst;
    }
    
    public List<List<String>> read(InputStream inputStream, boolean isExcel2003, int validColumn)
    {
        List<List<String>> dataLst = null;
        try
        {
            Workbook wb = null;
            if (isExcel2003)
            {
                wb = new HSSFWorkbook(inputStream);
            }
            else
            {
                wb = new XSSFWorkbook(inputStream);
            }
            dataLst = read(wb, validColumn);
        }
        catch (IOException e)
        {
        }
        catch (Exception e)
        {

        }
        
        return dataLst;
        
    }
    
    private List<List<String>> read(Workbook wb, int validColumn)
    {
        
        List<List<String>> dataLst = new ArrayList<List<String>>();
        
        Sheet sheet = wb.getSheetAt(0);
        long column = sheet.getRow(0).getPhysicalNumberOfCells();
        if(validColumn != column){
        }
        
        this.totalRows = sheet.getPhysicalNumberOfRows();
        
        if (this.totalRows >= 1 && sheet.getRow(0) != null)
        {
            this.totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }
        for (int r = 0; r < this.totalRows; r++)
        {
            Row row = sheet.getRow(r);
            
            if (row == null)
            {
                continue;
            }
            List<String> rowLst = new ArrayList<String>();
            
            for (int c = 0; c < this.getTotalCells(); c++)
            {
                Cell cell = row.getCell(c);
                String cellValue = "";
                if (null != cell)
                {
                    
                    switch (cell.getCellType())
                    {
                        case HSSFCell.CELL_TYPE_NUMERIC:
                            BigDecimal bd = new BigDecimal(
                                    cell.getNumericCellValue());
                            cellValue = bd.toPlainString();
                            break;
                        case HSSFCell.CELL_TYPE_STRING:
                            cellValue = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_FORMULA:
                            cellValue = cell.getCellFormula();
                            break;
                        case HSSFCell.CELL_TYPE_BLANK:
                            cellValue = "";
                            break;
                        case HSSFCell.CELL_TYPE_ERROR:
                            cellValue = "";
                            break;
                        default:
                            cellValue = "";
                            break;
                    }
                }
                rowLst.add(cellValue);
            }
            dataLst.add(rowLst);
        }
        return dataLst;
        
    }
    
    public boolean isExcel2003(String filePath)
    {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }
    
    public boolean isExcel2007(String filePath)
    {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

}
