package jaemisseo.man.reportman.handler

import jaemisseo.man.reportman.model.ColumnOption
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.usermodel.XSSFCell

class RowWriteHandler {

    void setHeaderData(XSSFCell cell, Object data){
        cell.setCellValue(data)
    }

    void setHeaderStyle(XSSFCell cell, Object data, CellStyle cellStyle){
        cell.setCellStyle(cellStyle)
    }

    void setData(XSSFCell cell, Object data){
        cell.setCellValue(data)
    }

    void setStyle(XSSFCell cell, Object data, CellStyle cellStyle){
        cell.setCellStyle(cellStyle)
    }

    boolean checkTwoTone(Map<String, ColumnOption> columnOptMap, Map<?,?> pkMap, Object rowData){
        boolean isTwoToneChange = false
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            if ( !pkMap || (pkMap.containsKey(attr) && pkMap[attr] != rowData[attr]) ){
                pkMap[attr] = rowData[attr]
                isTwoToneChange = true
            }
        }
        return isTwoToneChange
    }



}
