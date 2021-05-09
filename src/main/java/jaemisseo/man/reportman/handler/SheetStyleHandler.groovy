package jaemisseo.man.reportman.handler

import jaemisseo.man.ReportMan
import jaemisseo.man.reportman.model.CalculatedSheetInfo
import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.ReportSheetInfo
import jaemisseo.man.reportman.model.SheetOption
import jaemisseo.man.reportman.model.StyleOption
import jaemisseo.man.reportman.util.ReportUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.SheetConditionalFormatting
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet

class SheetStyleHandler {

    void setDefaultSize(XSSFSheet sheet, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        if (sheetOpt.width > -1)
            sheet.setDefaultColumnWidth(sheetOpt.width)
        if (sheetOpt.height > -1)
            sheet.setDefaultRowHeight((short)sheetOpt.height)
    }

    void setColumnWidth(XSSFSheet sheet, int index, int width){
        if (width > -1)
            sheet.setColumnWidth(index, width)
    }

    void setRowHeight(XSSFRow row, int height){
        if (height > -1)
            row.setHeight((short)height)
    }




    void autoFilter(Sheet sheet, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        if (sheetOpt.autoFilter){
            int headerLastIndex  = csi.headerLastIndex
            int headerStartX = csi.headerStartX
            int headerStartY = csi.headerStartY
            int dataStartX = csi.dataStartX
            int dataStartY = csi.dataStartY
            int dataSize = csi.dataSize
            int[] range = (sheetOpt.autoFilter == ReportMan.RANGE_AUTO) ? [headerStartX, headerStartY, dataStartX + headerLastIndex, dataStartY + dataSize] : ReportUtil.getRangeNumberIndexArray(sheetOpt.autoFilter)
            sheet.setAutoFilter(new CellRangeAddress(range[1], range[3], range[0], range[2]))
        }
    }

    void freezePane(Sheet sheet, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        if (sheetOpt.freezePane){
            int headerLastIndex  = csi.headerLastIndex
            int headerStartX = csi.headerStartX
            int headerStartY = csi.headerStartY
            int dataStartX = csi.dataStartX
            int dataStartY = csi.dataStartY
            int[] range = (sheetOpt.freezePane == ReportMan.RANGE_AUTO) ? [headerStartX, headerStartY +1] : ReportUtil.getRangeNumberIndexArray(sheetOpt.freezePane)
            if (range){
                if (range.size() == 4 && range[2] != null && range[3] != null)
                    sheet.createFreezePane(range[0], range[1], range[2], range[3])
                else
                    sheet.createFreezePane(range[0], range[1])
            }
        }
    }

    void sheetHighlightStyle(Sheet sheet, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        if (sheetOpt.highlightStyle){
            StyleOption styleOpt = sheetOpt.highlightStyle
            String condition = ReportUtil.parseVariable(styleOpt.option.condition, sheetOpt.replaceMap)
            String range = (styleOpt.option.range == ReportMan.RANGE_AUTO) ? ReportUtil.getResolvedRange(ReportMan.RANGE_DATA_ALL, 0, sheetInfo) : ReportUtil.getResolvedRange(styleOpt.option.range, sheetOpt, 0)
            ReportUtil.addFormattingRule(sheet, condition, range, styleOpt)
        }
    }


}