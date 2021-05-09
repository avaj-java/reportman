package jaemisseo.man.reportman.util

import jaemisseo.man.ReportMan
import jaemisseo.man.reportman.model.CalculatedSheetInfo
import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.ReportSheetInfo
import jaemisseo.man.reportman.model.SheetOption
import jaemisseo.man.reportman.model.StyleOption
import org.apache.poi.ss.usermodel.BorderFormatting
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.FontFormatting
import org.apache.poi.ss.usermodel.PatternFormatting
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.SheetConditionalFormatting
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.util.regex.Matcher
import java.util.regex.Pattern

class ReportUtil {

    static StyleOption mergeStyleOption(StyleOption a, StyleOption b){
        StyleOption mergedStyleOption = new StyleOption()
        boolean isMerged = false
        if (a && b){
            if (b.option.apply == ReportMan.APPLY_NEW){
                return b
            }
            mergedStyleOption.properties.keySet().each{ String attr ->
                if (attr != 'class' && attr != 'option'){
                    def attrA = a[attr]
                    def attrB = b[attr]
                    if (attrA != attrB){
                        if ( (attrB instanceof Short && attrB > -1)
                                || (attrB instanceof Integer && attrB > -1)
                                || (attrB instanceof Boolean && attrB)){
                            mergedStyleOption[attr] = attrB
                            isMerged = true
                        }else{
                            mergedStyleOption[attr] = attrA
                        }
                    }else{
                        mergedStyleOption[attr] = attrA
                    }
                }
            }
        }
        if (isMerged)
            return mergedStyleOption
        else if (!a && !b)
            return null
        else if (!a && b)
            return b
        else
            return a
    }







    static XSSFCellStyle generateCellStyle(StyleOption styleOpt, XSSFWorkbook workbook){
        XSSFCellStyle cellStyle = workbook.createCellStyle()
        XSSFFont font = workbook.createFont()
        if (styleOpt){
            //CUSTOM STYLE
            if (styleOpt.border > -1){
                cellStyle.setBorderLeft(styleOpt.border)
                cellStyle.setBorderRight(styleOpt.border)
                cellStyle.setBorderBottom(styleOpt.border)
                cellStyle.setBorderTop(styleOpt.border)
            }
            if (styleOpt.foreground > -1){
                cellStyle.setFillForegroundColor(styleOpt.foreground)
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            }
            if (styleOpt.background > -1){
                cellStyle.setFillBackgroundColor(styleOpt.background)
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            }
            if (styleOpt.fontSize > -1){
                font.setFontHeightInPoints(styleOpt.fontSize)
                cellStyle.setFont(font)
            }
            //POI STYLE
            List styleList = ['alignment', 'verticalAlignment', 'fillForegroundColor', 'fillBackgroundColor', 'borderTop', 'borderBottom', 'borderLeft', 'borderRight', 'wrapText']
            List fontList = ['color', 'fontHeightInPoints', 'bold', 'italic']
            styleList.each{ String prop ->
                if ((styleOpt[prop] instanceof Short && styleOpt[prop] > -1)
                        || (styleOpt[prop] instanceof Integer && styleOpt[prop] > -1)
                        || (styleOpt[prop] instanceof Boolean)
                        || (styleOpt[prop] instanceof FillPatternType && styleOpt.fillPattern != FillPatternType.NO_FILL)){
                    cellStyle[prop]  = styleOpt[prop]
                }
            }
            fontList.each{ String prop ->
                if ((styleOpt[prop] instanceof Short && styleOpt[prop] > -1)
                        || (styleOpt[prop] instanceof Integer && styleOpt[prop] > -1)
                        || (styleOpt[prop] instanceof Boolean)){
                    font[prop] = styleOpt[prop]
                    cellStyle.setFont(font)
                }
            }
        }
        return cellStyle
    }











    static int[] getRangeNumberIndexArray(String range){
        List<Integer> rangeNumberIndexes = []
        if (range.contains(",")){
            String[] temp = range.split("\\s*,\\s*")
            temp.eachWithIndex{ String it, int i ->
                rangeNumberIndexes[i] = Integer.parseInt(it)
            }
        }else if (range.contains(":")){
            String[] temp = range.split("\\s*:\\s*")
            CellReference cellRefRangeStart = new CellReference(temp[0])
            CellReference cellRefRangeEnd = new CellReference(temp[1])
            rangeNumberIndexes[0] = cellRefRangeStart.getCol()
            rangeNumberIndexes[1] = cellRefRangeStart.getRow()
            rangeNumberIndexes[2] = cellRefRangeEnd.getCol()
            rangeNumberIndexes[3] = cellRefRangeEnd.getRow()
        }
        return rangeNumberIndexes.toArray() as int[]
    }

    static String getResolvedRange(String range, int index, ReportSheetInfo sheetInfo){
        if (!range)
            return null
        SheetOption so = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        if (range == ReportMan.RANGE_ALL){
            range = "${csi.headerStartX},${csi.headerStartY +1},${csi.headerStartX + csi.headerLastIndex},${csi.dataStartY + csi.dataSize}"
        }else if (range == ReportMan.RANGE_COLUMN){
            range = "${csi.headerStartX + index},${csi.headerStartY +1},${csi.headerStartX + index},${csi.dataStartY + csi.dataSize}"
        }else if (range == ReportMan.RANGE_HEADER_ALL){
            range = "${csi.headerStartX},${csi.headerStartY +1},${csi.headerStartX + csi.headerLastIndex},${csi.headerStartY}"
        }else if (range == ReportMan.RANGE_HEADER_COLUMN) {
            range = "${csi.headerStartX},${csi.headerStartY +1},${csi.headerStartX},${csi.headerStartY +1}"
        }else if (range == ReportMan.RANGE_DATA_ALL){
            range = "${csi.dataStartX},${csi.dataStartY +1},${csi.dataStartX + csi.headerLastIndex},${csi.dataStartY + csi.dataSize}"
        }else if (range == ReportMan.RANGE_DATA_COLUMN){
            range = "${csi.dataStartX + index},${csi.dataStartY +1},${csi.dataStartX + index},${csi.dataStartY + csi.dataSize}"
        }
        return getRangeColumnString(range)
    }

    static String getColumnString(int columnIndex){
        return CellReference.convertNumToColString(columnIndex)
    }

    static int getColumnIndex(String columnString){
        return CellReference.convertColStringToIndex(columnString)
    }



    static String getRangeColumnString(int[] numberIndexArray){
        String rangeColumnString
        String colStart
        String rowStart
        String colEnd
        String rowEnd
        if (numberIndexArray){
            colStart = getColumnString(numberIndexArray[0])
            rowStart = "${numberIndexArray[1]}"
            if (numberIndexArray.size() >= 4){
                colEnd = getColumnString(numberIndexArray[2])
                rowEnd = "${numberIndexArray[3]}"
            }
            rangeColumnString = "${colStart}${rowStart}:${(colEnd)?:colStart}${(rowEnd)?:rowStart}"

        }
        return rangeColumnString
    }

    static String getRangeColumnString(String range){
        String rangeColumnString = ""
        if (range.contains(",")){
            int[] numberIndexArray = getRangeNumberIndexArray(range)
            rangeColumnString = getRangeColumnString(numberIndexArray)
        }else if (range.contains(":")){
        }
        return rangeColumnString
    }





    static void setFormattingStyle(XSSFConditionalFormattingRule conditionalFormatRule, StyleOption styleOpt){
        PatternFormatting pf = conditionalFormatRule.createPatternFormatting()
        FontFormatting ff = conditionalFormatRule.createFontFormatting()
        BorderFormatting bf = conditionalFormatRule.createBorderFormatting()
        //CUSTOM STYLE
        if (styleOpt.border > -1){
            bf.setBorderLeft(styleOpt.border)
            bf.setBorderRight(styleOpt.border)
            bf.setBorderBottom(styleOpt.border)
            bf.setBorderTop(styleOpt.border)
        }
        if (styleOpt.background > -1){
            pf.setFillBackgroundColor(styleOpt.background)
            pf.setFillPattern(PatternFormatting.SOLID_FOREGROUND)
        }
        if (styleOpt.foreground > -1){
            pf.setFillForegroundColor(styleOpt.foreground)
            pf.setFillPattern(PatternFormatting.SOLID_FOREGROUND)
        }
        //POI's STYLE
        if (styleOpt.fillBackgroundColor && styleOpt.fillBackgroundColor > -1)
            pf.setFillBackgroundColor(styleOpt.fillBackgroundColor)
        if (styleOpt.fillForegroundColor && styleOpt.fillForegroundColor > -1)
            pf.setFillForegroundColor(styleOpt.fillForegroundColor)
        if (styleOpt.fillPattern && styleOpt.fillPattern != FillPatternType.NO_FILL)
            pf.setFillPattern(styleOpt.fillPattern)
        if (styleOpt.color && styleOpt.color > -1)
            ff.setFontColorIndex(styleOpt.color)
        if (styleOpt.borderTop && styleOpt.borderTop > -1)
            bf.setBorderTop(styleOpt.borderTop)
        if (styleOpt.borderBottom && styleOpt.borderBottom > -1)
            bf.setBorderBottom(styleOpt.borderBottom)
        if (styleOpt.borderLeft && styleOpt.borderLeft > -1)
            bf.setBorderLeft(styleOpt.borderLeft)
        if (styleOpt.borderRight && styleOpt.borderRight > -1)
            bf.setBorderRight(styleOpt.borderRight)
    }



    static CalculatedSheetInfo calculatePositionData(ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.makeCalculatedSheetInfo()

        //dataStartY
        boolean hasHd = false
        int hLastIndex = 0
        int hStartX = 0
        int hStartY = 0
        int dStartX = 0
        int dStartY = 0
        if (sheetOpt.headerPosition){
            int[] pos = getRangeNumberIndexArray(sheetOpt.headerPosition)
            hStartX = pos[0]?:0
            hStartY = pos[1]?:0
        }
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            int index = cOpt.index
            String headerName = cOpt.headerName
            if (headerName && index != -1){
                hasHd = true
                hLastIndex = (hLastIndex > index) ? hLastIndex : index
            }
        }
        if (sheetOpt.dataPosition){
            int[] pos = getRangeNumberIndexArray(sheetOpt.dataPosition)
            dStartX = pos[0] ?: 0
            dStartY = pos[1] ?: 0
        }else if (!hasHd){
            dStartX = 0
            dStartY = 0
        }else{
            dStartX = hStartX
            dStartY = hStartY + 1
        }
        csi.with{
            modeHeader = hasHd
            headerLastIndex = hLastIndex
            headerStartX = hStartX
            headerStartY = hStartY
            dataStartX = dStartX
            dataStartY = dStartY
        }
        return csi;
    }

    //TODO: 계산Data는 csi로 가져가자
    static void calculateStyle(XSSFWorkbook workbook, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        StyleOption mergedHeaderStyle = ReportUtil.mergeStyleOption(sheetOpt.style, sheetOpt.headerStyle)
        StyleOption mergedDataStyle = ReportUtil.mergeStyleOption(sheetOpt.style, sheetOpt.dataStyle)
        StyleOption mergedDataTwoToneStyle = ReportUtil.mergeStyleOption(sheetOpt.style, sheetOpt.dataTwoToneStyle)
        sheetOpt.sheetHeaderRowStyle = ReportUtil.generateCellStyle(mergedHeaderStyle, workbook)
        sheetOpt.sheetDataRowStyle = ReportUtil.generateCellStyle(mergedDataStyle, workbook)
        sheetOpt.sheetDataTwoToneRowStyle = ReportUtil.generateCellStyle(mergedDataTwoToneStyle, workbook)
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            StyleOption mergedHeaderColumnStyle = ReportUtil.mergeStyleOption(mergedHeaderStyle, cOpt.headerStyle)
            StyleOption mergedDataColumnStyle = ReportUtil.mergeStyleOption(mergedDataStyle, cOpt.dataStyle)
            StyleOption mergedDataTwoToneColumnStyle = ReportUtil.mergeStyleOption(mergedDataTwoToneStyle, cOpt.dataStyle)
            if (mergedHeaderColumnStyle && mergedHeaderColumnStyle == mergedHeaderStyle){
                sheetOpt.columnHeaderCellStyleMap[attr] = sheetOpt.sheetHeaderRowStyle
            }else{
                sheetOpt.columnHeaderCellStyleMap[attr] = ReportUtil.generateCellStyle(mergedHeaderColumnStyle, workbook)
            }
            if (mergedDataColumnStyle && mergedDataColumnStyle == mergedDataStyle){
                sheetOpt.columnDataCellStyleMap[attr] = sheetOpt.sheetDataRowStyle
            }else{
                sheetOpt.columnDataCellStyleMap[attr] = ReportUtil.generateCellStyle(mergedDataColumnStyle, workbook)
            }
            if (mergedDataTwoToneColumnStyle && mergedDataTwoToneColumnStyle == mergedDataTwoToneStyle){
                sheetOpt.columnDataTwoToneCellStyleMap[attr] = sheetOpt.sheetDataTwoToneRowStyle
            }else{
                sheetOpt.columnDataTwoToneCellStyleMap[attr] = ReportUtil.generateCellStyle(mergedDataTwoToneColumnStyle, workbook)
            }
        }
    }


    static Map<String, String> generateReplaceMap(ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        Map<String, String> replaceMap = [:]
        replaceMap.headerStartCol = getColumnString(csi.headerStartX)
        replaceMap.headerStartRow = csi.headerStartY +1
        replaceMap.dataStartCol = getColumnString(csi.dataStartX)
        replaceMap.dataStartRow = csi.dataStartY +1
        replaceMap.dataEndCol = getColumnString(csi.dataStartX + csi.headerLastIndex)
        replaceMap.dataSize = csi.dataSize
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            int colIndex = cOpt.index
            if (colIndex > -1){
                replaceMap["${cOpt.index}"] = getColumnString(csi.dataStartX + colIndex)
                replaceMap[attr] = getColumnString(csi.dataStartX + colIndex)
            }
        }
        return replaceMap
    }

    static String parseVariable(String condition, Map replaceMap){
        if (!condition)
            return null
        String patternToGetVariable = '[$][{][^{]*\\w+[^{]*[}]'     // If variable contains some word in ${} then convert to User Set Value or...
        String codeRule = condition
        ///// Get String In ${ } step by step
        codeRule = codeRule.trim()
        String resultStr = codeRule
        Matcher matchedList = Pattern.compile(patternToGetVariable).matcher(codeRule)
        matchedList.each { String oneVal ->
            // 1. get String in ${ }
            String content = oneVal.replaceFirst('[\$]', '').replaceFirst('\\{', '').replaceFirst('\\}', '')
            // 3. Replace One ${ }
            if (content != null) {
                String replacement = replaceMap[content]
                resultStr = resultStr.replaceFirst(patternToGetVariable, replacement)
            }
        }
        return resultStr
    }


    public static addFormattingRule(Sheet sheet, String condition, String range, StyleOption styleOpt){
        //CREATE CONDITION
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting()
        // ConditionalFormatRule
        XSSFConditionalFormattingRule conditionalFormatRule = sheetCF.createConditionalFormattingRule(condition)
        // Then => Add Style
        ReportUtil.setFormattingStyle(conditionalFormatRule, styleOpt)
        // Range
        CellRangeAddress[] regions = [CellRangeAddress.valueOf(range)]
        //ADD CONDITION
        sheetCF.addConditionalFormatting(regions, conditionalFormatRule)
        return this
    }

}
