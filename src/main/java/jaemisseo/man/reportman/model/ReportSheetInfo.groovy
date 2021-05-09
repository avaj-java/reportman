package jaemisseo.man.reportman.model

import org.apache.poi.ss.usermodel.Sheet

class ReportSheetInfo {

    Object mapperInstance

    SheetOption sheetOpt

    Map<String, ColumnOption> columnOptMap

    CalculatedSheetInfo calculatedSheetInfo

    Sheet sheet

    CalculatedSheetInfo makeCalculatedSheetInfo(){
        calculatedSheetInfo = calculatedSheetInfo ?: new CalculatedSheetInfo()
        return calculatedSheetInfo
    }

}
