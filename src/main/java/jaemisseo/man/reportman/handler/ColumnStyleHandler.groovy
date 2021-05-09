package jaemisseo.man.reportman.handler

import jaemisseo.man.ReportMan
import jaemisseo.man.reportman.model.CalculatedSheetInfo
import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.ReportSheetInfo
import jaemisseo.man.reportman.model.SheetOption
import jaemisseo.man.reportman.model.StyleOption
import jaemisseo.man.reportman.util.ReportUtil
import org.apache.poi.ss.usermodel.Sheet

class ColumnStyleHandler {

    void columnHighlightStyle(Sheet sheet, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        columnOptMap.each{ String attr, ColumnOption cOpt ->
            if (cOpt.highlightStyle){
                int index = cOpt.index
                StyleOption styleOpt = cOpt.highlightStyle
                String condition = ReportUtil.parseVariable(styleOpt.option.condition, csi.replaceMap)
                String range = (styleOpt.option.range == ReportMan.RANGE_AUTO) ? ReportUtil.getResolvedRange(ReportMan.RANGE_DATA_COLUMN, index, sheetInfo) : ReportUtil.getResolvedRange(styleOpt.option.range, index, sheetInfo)
                if (range)
                    ReportUtil.addFormattingRule(sheet, condition, range, styleOpt)
            }
        }
    }

}
