package jaemisseo.man

import jaemisseo.man.annotation.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.*
import org.slf4j.Logger
import org.slf4j.LoggerFactory

import java.lang.annotation.Annotation
import java.lang.reflect.Field
import java.util.regex.Matcher
import java.util.regex.Pattern

/**
 * Created by sujkim on 2017-02-05.
 */
class ReportMan {

    final Logger logger = LoggerFactory.getLogger(this.getClass());

    ReportMan(){}

    ReportMan(File file){
        this.outputStream = new FileOutputStream(file)
    }

    ReportMan(String fileName){
        this.outputStream = new FileOutputStream(new File(fileName))
    }

    ReportMan(OutputStream outputStream){
        this.outputStream = outputStream
    }



    public static final String RANGE_AUTO = "AUTO"
    public static final String RANGE_DATA_ALL = "DATA_ALL"
    public static final String RANGE_DATA_COLUMN = "DATA_COLUMN"
    public static final String RANGE_HEADER_ALL = "HEADER_ALL"
    public static final String RANGE_HEADER_COLUMN = "HEADER_COLUMN"
    public static final String RANGE_ALL = "ALL"
    public static final String RANGE_COLUMN = "COLUMN"

    public static final int APPLY_NEW = 1
    public static final int APPLY_MODIFY = 2

    public static final int HIGHLIGHT_CELL = 1
    public static final int HIGHLIGHT_ROW = 2
    public static final int HIGHLIGHT_HEADER = 3



    class SheetOption extends StyleOption{
        //SIZE
        int width
        int height
        int headerHeight
        //Position
        String headerPosition
        String dataPosition
        boolean hasHeader
        int headerLastIndex
        int headerStartX
        int headerStartY
        int dataStartX
        int dataStartY
        int dataSize
        Map replaceMap
        //
        String freezePane
        String autoFilter
        //Style Infomation
        StyleOption style
        StyleOption headerStyle
        StyleOption dataStyle
        StyleOption dataTwoToneStyle
        StyleOption highlightStyle
        //Default Style
        XSSFCellStyle sheetStyle
        XSSFCellStyle sheetHeaderRowStyle
        XSSFCellStyle sheetDataRowStyle
        XSSFCellStyle sheetDataTwoToneRowStyle
        //Combind Style
        Map<String, CellStyle> columnHeaderCellStyleMap = [:]
        Map<String, CellStyle> columnDataCellStyleMap = [:]
        Map<String, CellStyle> columnDataTwoToneCellStyleMap = [:]
    }

    class ColumnOption extends StyleOption{
        int index
        //SIZE
        int width
        int height
        //
        String headerName
        boolean isSheetNameField
        //
        StyleOption headerStyle
        StyleOption dataStyle
        StyleOption highlightStyle
    }

    class StyleOption{
        //POIs STYLE
        short color
        short fontHeightInPoints
        boolean bold
        boolean italic
        boolean wrapText
        short alignment
        short verticalAlignment
        short fillForegroundColor
        short fillBackgroundColor
        FillPatternType fillPattern
        short borderTop
        short borderBottom
        short borderLeft
        short borderRight
        //ReportMan CUSTOM PROEPRTIES
        short fontSize
        short border
        short foreground
        short background
        //OPTION
        Map option = [:]
    }

    OutputStream outputStream
    SheetOption sheetOpt
    Map<String, ColumnOption> columnOptMap
    Map<String, ColumnOption> additionalColumnOptionMap
    List<String> excludeColumnFieldNameList


    /**************************************************
     *
     * Setup
     *
     **************************************************/
    ReportMan setColumns(String columns){
        List<String> columnList = columns.split("\\s*,\\s*").toList()
        return setColumns(columnList)
    }

    ReportMan setColumns(String... column){
        return setColumns(column.toList())
    }

    ReportMan setColumns(List<String> columnList){
        columnList.each{ String columnString ->
            columnOptMap[columnString] = new ColumnOption(index: getColumnIndex(columnString))
        }
        return this
    }

    ReportMan setExcludeColumn(String... excludeColumnFieldNames){
        this.excludeColumnFieldNameList = excludeColumnFieldNames.toList()
        return this
    }

    ReportMan addAdditionalColumnOption(Map<String, ColumnOption> additionalColumnOptionMap){
        this.additionalColumnOptionMap = additionalColumnOptionMap
        return this
    }


    /**************************************************
     *
     * Read Excel File  ==> Data
     *
     **************************************************/
    List toRowList(InputStream inputStream){
        List allRowList = []
        Map<String, String> mapInstance = [:]
        excelToData(inputStream, mapInstance){ String sheetName, List rowList ->
            allRowList += rowList
        }
        return allRowList
    }

    List toRowList(InputStream inputStream, def instance){
        List allRowList = []
        excelToData(inputStream, instance){ String sheetName, List rowList ->
            allRowList += rowList
        }
        return allRowList
    }

    Map<String, List> toSheetMap(InputStream inputStream, def instance){
        Map sheetMap = [:]
        excelToData(inputStream, instance){ String sheetName, List rowList ->
            sheetMap[sheetName] = rowList
        }
        return sheetMap
    }

    private boolean excelToData(InputStream inputStream, def instance, Closure closure){
        sheetOpt = sheetOpt ?: generateSheetOption(instance)
        columnOptMap = columnOptMap ?: generateColumnOptionMap(instance)
        generatePositionData(sheetOpt, columnOptMap)
        int dataStartX = sheetOpt.dataStartX
        int dataStartY = sheetOpt.dataStartY
        // 1. Get excel file
        Workbook workbook = WorkbookFactory.create( inputStream )
        int sheetSize = workbook.numberOfSheets
        // 2. Get sheet.
        for (int sheetIndex=0; sheetIndex<sheetSize; sheetIndex++){
            // 3. Get rowList.
            Sheet sheet = workbook.getSheetAt(sheetIndex)
            String sheetName = sheet.getSheetName()
            List rowList = []
            int rowIndex = 0
            for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext();){
                // 4. Get Cells.
                Row row = rowsIT.next()
                if (rowIndex++ < dataStartY){
                    continue
                }
                rowList << getInstanceForOneRow(instance, row, dataStartX, sheetName)
            }
            closure(sheetName, rowList)
        }
        return true
    }

    def getInstanceForOneRow(def instance, Row row, Integer dataStartX, String sheetName){
        def rowDto = instance.getClass().newInstance()
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            // 5. Get Only Mapping Cell
            Cell cell = (cOpt.index > -1) ? row.getCell(dataStartX + cOpt.index) : null
            if (cell){
                if (attr){
                    Class clazz = rowDto.metaClass.properties.find{ it.name == attr }.type
                    cell.setCellType(Cell.CELL_TYPE_STRING)
                    if (clazz == Integer.class){
                        String value = cell.getStringCellValue()
                        rowDto[attr] = value.isNumber() ? ((Double)Double.parseDouble(value)).intValue() : null
                    }else{
                        rowDto[attr] = cell.getStringCellValue()
                    }
//                    println rowDto[attr]
                }
            }else if (cOpt.isSheetNameField){
                rowDto[attr] = sheetName
            }
        }
        return rowDto
    }



    /**************************************************
     * Write Excel File  <== Data
     *
     *    Example>
     *      public downloadExcel(HttpServletResponse res){
     *          List<User> users = getUserList();
     *          ServletOutputStream out = setHeaderForDownloadExcel(res, fileName).getOutputStream();
     *          new ReportMan(out).write(users);
     *      }
     *
     *      private HttpServletResponse setHeaderForDownloadExcel(HttpServletResponse res, String fileName) {
     *         res.setHeader("Content-disposition","attachment;filename=" +fileName);
     *         res.setHeader("Content-Type", "application/vnd.ms-excel; charset=MS949");
     *         res.setHeader("Content-Description", "JSP Generated Data");
     *         res.setHeader("Content-Transfer-Encoding", "binary;");
     *         res.setHeader("Pragma", "no-cache;");
     *         res.setHeader("Expires", "-1;");
     *         return res;
     *     }
     **************************************************/
    boolean write(Class clazz){
        List dummyList = [clazz.newInstance()]
        return write(["Sheet1":dummyList])
    }

    boolean write(List allRowList){
        return write(["Sheet1":allRowList])
    }

    boolean write(String sheetName, List allRowList){
        return write(["${sheetName}":allRowList])
    }

    boolean write(List allRowList, String sheetFieldName){
        Map sheetMap = [:]
        allRowList.each{
            List rowList = sheetMap[it[sheetFieldName]]
            if (!rowList)
                sheetMap[it[sheetFieldName]] = [it]
            else
                rowList << it
        }
        return write(sheetMap)
    }

    boolean write(Map sheetMap){
        //Recognize Instance Annotation
        for (String key : sheetMap.keySet()){
            List rowList = sheetMap[key]
            for (def instance : rowList) {
                sheetOpt = sheetOpt ?: generateSheetOption(instance)
                columnOptMap = columnOptMap ?: generateColumnOptionMap(instance)
                generatePositionData(sheetOpt, columnOptMap)
                break
            }
        }
        //Create Excel
        XSSFWorkbook workbook = new XSSFWorkbook();
        //Check SheetName (null => No Sheet Name)
        sheetMap.findAll{ String sheetName, List dataRowList ->
            return !sheetName
        }.each{ String sheetName, List dataRowList ->
            if (!sheetMap['No Sheet Name'])
                sheetMap['No Sheet Name'] = []
            (sheetMap['No Sheet Name'] as List).addAll(dataRowList)
        }
        sheetMap.remove(null)
        //Create Sheet
        sheetMap.each{ String sheetName, List dataRowList ->
            XSSFSheet sheet = workbook.createSheet(sheetName)
            Map cellStyleMap
            boolean hasHeader = sheetOpt.hasHeader
            int headerLastIndex  = sheetOpt.headerLastIndex
            int headerStartX = sheetOpt.headerStartX
            int headerStartY = sheetOpt.headerStartY
            int dataStartX = sheetOpt.dataStartX
            int dataStartY = sheetOpt.dataStartY
            int dataSize = dataRowList.size()

            //Generate CellStyle
            StyleOption mergedHeaderStyle = mergeStyleOption(sheetOpt.style, sheetOpt.headerStyle)
            StyleOption mergedDataStyle = mergeStyleOption(sheetOpt.style, sheetOpt.dataStyle)
            StyleOption mergedDataTwoToneStyle = mergeStyleOption(sheetOpt.style, sheetOpt.dataTwoToneStyle)
            sheetOpt.sheetHeaderRowStyle = generateCellStyle(mergedHeaderStyle, workbook)
            sheetOpt.sheetDataRowStyle = generateCellStyle(mergedDataStyle, workbook)
            sheetOpt.sheetDataTwoToneRowStyle = generateCellStyle(mergedDataTwoToneStyle, workbook)
            columnOptMap.each{ String attr, ColumnOption cOpt ->
                StyleOption mergedHeaderColumnStyle = mergeStyleOption(mergedHeaderStyle, cOpt.headerStyle)
                StyleOption mergedDataColumnStyle = mergeStyleOption(mergedDataStyle, cOpt.dataStyle)
                StyleOption mergedDataTwoToneColumnStyle = mergeStyleOption(mergedDataTwoToneStyle, cOpt.dataStyle)
                if (mergedHeaderColumnStyle && mergedHeaderColumnStyle == mergedHeaderStyle){
                    sheetOpt.columnHeaderCellStyleMap[attr] = sheetOpt.sheetHeaderRowStyle
                }else{
                    sheetOpt.columnHeaderCellStyleMap[attr] = generateCellStyle(mergedHeaderColumnStyle, workbook)
                }
                if (mergedDataColumnStyle && mergedDataColumnStyle == mergedDataStyle){
                    sheetOpt.columnDataCellStyleMap[attr] = sheetOpt.sheetDataRowStyle
                }else{
                    sheetOpt.columnDataCellStyleMap[attr] = generateCellStyle(mergedDataColumnStyle, workbook)
                }
                if (mergedDataTwoToneColumnStyle && mergedDataTwoToneColumnStyle == mergedDataTwoToneStyle){
                    sheetOpt.columnDataTwoToneCellStyleMap[attr] = sheetOpt.sheetDataTwoToneRowStyle
                }else{
                    sheetOpt.columnDataTwoToneCellStyleMap[attr] = generateCellStyle(mergedDataTwoToneColumnStyle, workbook)
                }
            }

            //SIZE(Sheet) - "왜 width는 안먹냐"
            setDefaultSize(sheetOpt, sheet)

            //Create Row(Header)
            XSSFRow headerRow
            //CellStyleMap
            cellStyleMap = sheetOpt.columnHeaderCellStyleMap
            //Row(Header)
            if (hasHeader){
                headerRow = sheet.createRow(headerStartY)
                columnOptMap.each{ String attr, ColumnOption cOpt ->
                    int index = cOpt.index
                    String headerName = cOpt.headerName
                    if (headerName && index != -1){
                        XSSFCell cell = headerRow.createCell(headerStartX + index)
                        //VALUE(Header)
                        cell.setCellValue(headerName)
                        //SIZE - WIDTH
                        setColumnWidth(sheet, headerStartX + index, cOpt.width)
                        //STYLE(Header)
                        cell.setCellStyle(cellStyleMap[attr])
                    }
                }
                //SIZE - HEIGHT
                setRowHeight(headerRow, sheetOpt.headerHeight)
            }


            //Create Row(Data)
            //Option - DataTwoToneStyle
            boolean modeTowToneChange = (!!sheetOpt.dataTwoToneStyle)
            boolean modeTwoToneToggleChange
            List<String> pkList
            Map pkMap = [:]
            if (modeTowToneChange){
                pkList = sheetOpt.dataTwoToneStyle.option.pk.split("\\s*,\\s*")
                if (pkList[0])
                    pkList.each{ pkMap[it] = null }
                else
                    modeTwoToneToggleChange = true
            }
            //CellStyleMap
            cellStyleMap = (modeTowToneChange) ? sheetOpt.columnDataTwoToneCellStyleMap : sheetOpt.columnDataCellStyleMap
            //Row(Data)
            dataRowList.eachWithIndex{ rowData, rowIndex ->
                XSSFRow row = sheet.createRow(rowIndex + dataStartY)
                //Check DataTwoToneChange
                if (modeTowToneChange){
                    boolean isTwoToneChange = false
                    columnOptMap.each{ String attr, ColumnOption cOpt ->
                        if ( !pkMap || (pkMap.containsKey(attr) && pkMap[attr] != rowData[attr]) ){
                            pkMap[attr] = rowData[attr]
                            isTwoToneChange = true
                        }
                    }
                    if (isTwoToneChange || modeTwoToneToggleChange)
                        cellStyleMap = (cellStyleMap == sheetOpt.columnDataCellStyleMap) ? sheetOpt.columnDataTwoToneCellStyleMap : sheetOpt.columnDataCellStyleMap
                }
                //Create Cell
                columnOptMap.each{ String attr, ColumnOption cOpt ->
                    int index = cOpt.index
                    if (index != -1){
                        XSSFCell cell = row.createCell(index + dataStartX);
                        //VALUE
                        cell.setCellValue(rowData[attr])
                        //STYLE(Data)
                        cell.setCellStyle(cellStyleMap[attr])
                    }
                }
            }

            //Set Position Info
            sheetOpt.dataSize = dataSize
            sheetOpt.replaceMap = generateReplaceMap(sheetOpt)

            //AutoFilter
            if (sheetOpt.autoFilter){
                int[] range = (sheetOpt.autoFilter == ReportMan.RANGE_AUTO) ? [headerStartX, headerStartY, dataStartX + headerLastIndex, dataStartY + dataSize] : getRangeNumberIndexArray(sheetOpt.autoFilter)
                sheet.setAutoFilter(new CellRangeAddress(range[1], range[3], range[0], range[2]))
            }

            //FreezePane
            if (sheetOpt.freezePane){
                int[] range = (sheetOpt.freezePane == ReportMan.RANGE_AUTO) ? [headerStartX, headerStartY +1] : getRangeNumberIndexArray(sheetOpt.freezePane)
                if (range){
                    if (range.size() == 4 && range[2] != null && range[3] != null)
                        sheet.createFreezePane(range[0], range[1], range[2], range[3])
                    else
                        sheet.createFreezePane(range[0], range[1])
                }
            }

            //ConditionalFormattingRule - Sheet
            if (sheetOpt.highlightStyle){
                StyleOption styleOpt = sheetOpt.highlightStyle
                String condition = parseVariable(styleOpt.option.condition, sheetOpt.replaceMap)
                String range = (styleOpt.option.range == ReportMan.RANGE_AUTO) ? getResolvedRange(ReportMan.RANGE_DATA_ALL, sheetOpt, 0) : getResolvedRange(styleOpt.option.range, sheetOpt, 0)
                addFormattingRule(sheet, condition, range, styleOpt)
            }
            //ConditionalFormattingRule - Column
            columnOptMap.each{ String attr, ColumnOption cOpt ->
                if (cOpt.highlightStyle){
                    int index = cOpt.index
                    StyleOption styleOpt = cOpt.highlightStyle
                    String condition = parseVariable(styleOpt.option.condition, sheetOpt.replaceMap)
                    String range = (styleOpt.option.range == ReportMan.RANGE_AUTO) ? getResolvedRange(ReportMan.RANGE_DATA_COLUMN, sheetOpt, index) : getResolvedRange(styleOpt.option.range, sheetOpt, index)
                    addFormattingRule(sheet, condition, range, styleOpt)
                }
            }

        }

        //Write Excel File
        workbook.write(outputStream)
        outputStream.flush()
        outputStream.close()
        return true
    }





    ReportMan setDefaultSize(SheetOption sOpt, XSSFSheet sheet){
        if (sOpt.width > -1)
            sheet.setDefaultColumnWidth(sOpt.width)
        if (sOpt.height > -1)
            sheet.setDefaultRowHeight((short)sOpt.height)
        return this
    }

    ReportMan setColumnWidth(XSSFSheet sheet, int index, int width){
        if (width > -1)
            sheet.setColumnWidth(index, width)
        return this
    }
    ReportMan setRowHeight(XSSFRow row, int height){
        if (height > -1)
            row.setHeight((short)height)
        return this
    }





    StyleOption mergeStyleOption(StyleOption a, StyleOption b){
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

    XSSFCellStyle generateCellStyle(StyleOption styleOpt, XSSFWorkbook workbook){
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





    SheetOption generateSheetOption(def instance){
        SheetOption sheet
        if (instance.getClass().getAnnotation(ReportSheet.class)){
            ReportSheet sheetAnt = instance.getClass().getAnnotation(ReportSheet.class)
            ReportSheetStyle allAnt = instance.getClass().getAnnotation(ReportSheetStyle.class)
            ReportSheetHeaderStyle headerAnt = instance.getClass().getAnnotation(ReportSheetHeaderStyle.class)
            ReportSheetDataStyle dataAnt = instance.getClass().getAnnotation(ReportSheetDataStyle.class)
            ReportSheetDataTwoToneStyle dataTwoToneAnt = instance.getClass().getAnnotation(ReportSheetDataTwoToneStyle.class)
            ReportSheetHighlightStyle highlightAnt = instance.getClass().getAnnotation(ReportSheetHighlightStyle.class)
            if (sheetAnt){
                sheet = new SheetOption(
                    width: sheetAnt.width(),
                    height: sheetAnt.height(),
                    headerHeight: sheetAnt.headerHeight(),
                    headerPosition: sheetAnt.headerPosition(),
                    dataPosition: sheetAnt.dataPosition(),
                    freezePane: sheetAnt.freezePane(),
                    autoFilter: sheetAnt.autoFilter(),
                )
                if (allAnt)
                    sheet.style = generateStyleOption(allAnt)
                if (headerAnt)
                    sheet.headerStyle = generateStyleOption(headerAnt)
                if (dataAnt)
                    sheet.dataStyle = generateStyleOption(dataAnt)
                if (dataTwoToneAnt){
                    sheet.dataTwoToneStyle = generateStyleOption(dataTwoToneAnt)
                    sheet.dataTwoToneStyle.option['pk'] = dataTwoToneAnt.pk()
                }
                if (highlightAnt)
                    sheet.highlightStyle = generateHightlightStyleOption(highlightAnt)
            }
        }
        return sheet
    }

    Map generateColumnOptionMap(def instance){
        Map<String, ColumnOption> columnOptMap = [:]
        instance.getClass().getDeclaredFields().each{ Field field ->
            ReportColumn columnAnt = field.getAnnotation(ReportColumn.class)
            ReportSheetName sheetNameAnt = field.getAnnotation(ReportSheetName.class)
            ReportColumnHeaderStyle headerAnt = field.getAnnotation(ReportColumnHeaderStyle.class)
            ReportColumnDataStyle dataAnt = field.getAnnotation(ReportColumnDataStyle.class)
            ReportColumnHighlightStyle highlightAnt = field.getAnnotation(ReportColumnHighlightStyle.class)
            if (columnAnt || sheetNameAnt){
                field.accessible = true
                String fieldName = field.name
                //- Checking Exclude Column
                if (!excludeColumnFieldNameList?.contains(fieldName)){
                    columnOptMap[fieldName] = new ColumnOption(
                            index: columnAnt ? columnAnt.index() : -1,
                            width: columnAnt ? columnAnt.width() : -1,
                            headerName: columnAnt ? columnAnt.headerName() : null,
                            isSheetNameField: sheetNameAnt ? true : false
                    )
                    if (headerAnt)
                        columnOptMap[fieldName].headerStyle = generateStyleOption(headerAnt)
                    if (dataAnt)
                        columnOptMap[fieldName].dataStyle = generateStyleOption(dataAnt)
                    if (highlightAnt)
                        columnOptMap[fieldName].highlightStyle = generateHightlightStyleOption(highlightAnt)
                }
            }
        }
        //Setup Additional Column Options
        additionalColumnOptionMap?.each{ String fieldName, ColumnOption columnOption ->
            columnOptMap[fieldName] = columnOption
        }
        //Sorting
        columnOptMap.sort{ a, b ->
            a.value.index <=> b.value.index
        }
        return columnOptMap
    }

    StyleOption generateStyleOption(Annotation ant){
        StyleOption styleOpt = new StyleOption(
                //POI STYLE
                color               : ant.color(),
                fontHeightInPoints  : (short)ant.fontHeightInPoints(),
                bold                : ant.bold(),
                italic              : ant.italic(),
                wrapText            : ant.wrapText(),
                alignment           : ant.alignment(),
                verticalAlignment   : ant.verticalAlignment(),
                fillForegroundColor : ant.fillForegroundColor(),
                fillBackgroundColor : ant.fillBackgroundColor(),
                fillPattern         : ant.fillPattern(),
                borderTop           : ant.borderTop(),
                borderBottom        : ant.borderBottom(),
                borderLeft          : ant.borderLeft(),
                borderRight         : ant.borderRight(),
                //ReportMan CUSTOM STYLE
                fontSize            : (short)ant.fontSize(),
                border              : ant.border(),
                foreground          : ant.foreground(),
                background          : ant.background()
        )
        styleOpt.option['apply'] = ant.apply()
        return styleOpt
    }

    StyleOption generateHightlightStyleOption(Annotation ant){
        StyleOption styleOpt = new StyleOption(
                //POI STYLE
                color               : ant.color(),
                fillForegroundColor : ant.fillForegroundColor(),
                fillBackgroundColor : ant.fillBackgroundColor(),
                fillPattern         : ant.fillPattern(),
                borderTop           : ant.borderTop(),
                borderBottom        : ant.borderBottom(),
                borderLeft          : ant.borderLeft(),
                borderRight         : ant.borderRight(),
                //ReportMan CUSTOM STYLE
                border              : ant.border(),
                foreground          : ant.foreground(),
                background          : ant.background()
        )
        styleOpt.option['condition'] = ant.condition()
        styleOpt.option['range'] = ant.range()
        return styleOpt
    }





    int[] getRangeNumberIndexArray(String range){
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

    String getResolvedRange(String range, SheetOption so, int index){
        if (range == ReportMan.RANGE_ALL){
            range = "${so.headerStartX},${so.headerStartY +1},${so.headerStartX + so.headerLastIndex},${so.dataStartY + so.dataSize}"
        }else if (range == ReportMan.RANGE_COLUMN){
            range = "${so.headerStartX + index},${so.headerStartY +1},${so.headerStartX + index},${so.dataStartY + so.dataSize}"
        }else if (range == ReportMan.RANGE_HEADER_ALL){
            range = "${so.headerStartX},${so.headerStartY +1},${so.headerStartX + so.headerLastIndex},${so.headerStartY}"
        }else if (range == ReportMan.RANGE_HEADER_COLUMN) {
            range = "${so.headerStartX},${so.headerStartY +1},${so.headerStartX},${so.headerStartY +1}"
        }else if (range == ReportMan.RANGE_DATA_ALL){
            range = "${so.dataStartX},${so.dataStartY +1},${so.dataStartX + so.headerLastIndex},${so.dataStartY + so.dataSize}"
        }else if (range == ReportMan.RANGE_DATA_COLUMN){
            range = "${so.dataStartX + index},${so.dataStartY +1},${so.dataStartX + index},${so.dataStartY + so.dataSize}"
        }
        return getRangeColumnString(range)
    }

    String getColumnString(int columnIndex){
        return CellReference.convertNumToColString(columnIndex)
    }

    int getColumnIndex(String columnString){
        return CellReference.convertColStringToIndex(columnString)
    }



    String getRangeColumnString(int[] numberIndexArray){
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

    String getRangeColumnString(String range){
        String rangeColumnString = ""
        if (range.contains(",")){
            int[] numberIndexArray = getRangeNumberIndexArray(range)
            rangeColumnString = getRangeColumnString(numberIndexArray)
        }else if (range.contains(":")){
        }
        return rangeColumnString
    }





    ReportMan addFormattingRule(Sheet sheet, String condition, String range, StyleOption styleOpt){
        //CREATE CONDITION
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting()
        // ConditionalFormatRule
        XSSFConditionalFormattingRule conditionalFormatRule = sheetCF.createConditionalFormattingRule(condition)
        // Then => Add Style
        setFormattingStyle(conditionalFormatRule, styleOpt)
        // Range
        CellRangeAddress[] regions = [CellRangeAddress.valueOf(range)]
        //ADD CONDITION
        sheetCF.addConditionalFormatting(regions, conditionalFormatRule)
        return this
    }

    private void setFormattingStyle(XSSFConditionalFormattingRule conditionalFormatRule, StyleOption styleOpt){
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



    private void generatePositionData(SheetOption so, Map<String, ColumnOption> coMap){
        //dataStartY
        boolean hasHd = false
        int hLastIndex = 0
        int hStartX = 0
        int hStartY = 0
        int dStartX = 0
        int dStartY = 0
        if (so.headerPosition){
            int[] pos = getRangeNumberIndexArray(so.headerPosition)
            hStartX = pos[0]?:0
            hStartY = pos[1]?:0
        }
        coMap.each{ String attr, ColumnOption cOpt ->
            int index = cOpt.index
            String headerName = cOpt.headerName
            if (headerName && index != -1){
                hasHd = true
                hLastIndex = (hLastIndex > index) ? hLastIndex : index
            }
        }
        if (so.dataPosition){
            int[] pos = getRangeNumberIndexArray(so.dataPosition)
            dStartX = pos[0] ?: 0
            dStartY = pos[1] ?: 0
        }else if (!hasHd){
            dStartX = 0
            dStartY = 0
        }else{
            dStartX = hStartX
            dStartY = hStartY + 1
        }
        so.with{
            hasHeader = hasHd
            headerLastIndex = hLastIndex
            headerStartX = hStartX
            headerStartY = hStartY
            dataStartX = dStartX
            dataStartY = dStartY
        }
    }


    Map<String, String> generateReplaceMap(SheetOption so){
        Map<String, String> replaceMap = [:]
        replaceMap.headerStartCol = getColumnString(so.headerStartX)
        replaceMap.headerStartRow = so.headerStartY +1
        replaceMap.dataStartCol = getColumnString(so.dataStartX)
        replaceMap.dataStartRow = so.dataStartY +1
        replaceMap.dataEndCol = getColumnString(so.dataStartX + so.headerLastIndex)
        replaceMap.dataSize = so.dataSize
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            int colIndex = cOpt.index
            if (colIndex > -1){
                replaceMap["${cOpt.index}"] = getColumnString(so.dataStartX + colIndex)
                replaceMap[attr] = getColumnString(so.dataStartX + colIndex)
            }
        }
        return replaceMap
    }

    String parseVariable(String condition, Map replaceMap){
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


}
