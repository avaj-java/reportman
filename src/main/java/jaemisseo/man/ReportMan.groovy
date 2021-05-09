package jaemisseo.man

import jaemisseo.man.reportman.handler.ColumnStyleHandler
import jaemisseo.man.reportman.handler.RowWriteHandler
import jaemisseo.man.reportman.handler.SheetStyleHandler
import jaemisseo.man.reportman.model.CalculatedSheetInfo
import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.ReportOption
import jaemisseo.man.reportman.model.ReportSheetInfo
import jaemisseo.man.reportman.model.ReportSystem
import jaemisseo.man.reportman.model.SheetOption
import jaemisseo.man.reportman.util.AnnotationExtractor
import jaemisseo.man.reportman.util.ReportUtil
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.slf4j.Logger
import org.slf4j.LoggerFactory

class ReportMan {

    final Logger logger = LoggerFactory.getLogger(this.getClass());

    ReportMan(){
        init()
    }

    ReportMan(File file){
        init()
        if (!system.option)
            system.option = new ReportOption()
        system.option.file = file
    }

    ReportMan(String fileName){
        init()
        if (!system.option)
            system.option = new ReportOption()
        system.option.fileName = fileName
    }

    ReportMan(OutputStream outputStream){
        init()
        if (!system.option)
            system.option = new ReportOption()
        system.option.outputStream = outputStream
    }

    ReportMan(ReportOption reportOption){
        init()
        setForceOption(reportOption)
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

//    ReportOption option
    ReportSystem system
    SheetStyleHandler sheetStyleHandler
    ColumnStyleHandler columnStyleHandler
    RowWriteHandler rowWriteHandler


    private init(){
//        this.option = new ReportOption()
        this.system = new ReportSystem()
        this.sheetStyleHandler = new SheetStyleHandler()
        this.columnStyleHandler = new ColumnStyleHandler()
        this.rowWriteHandler = new RowWriteHandler()
    }


    /**************************************************
     *
     * Setup
     *
     **************************************************/
    ReportMan setForceOption(ReportOption reportOption){
        system.option = reportOption
        return this
    }

    ReportMan setColumns(String columns){
        List<String> columnList = columns.split("\\s*,\\s*").toList()
        return setColumns(columnList)
    }

    ReportMan setColumns(String... column){
        return setColumns(column.toList())
    }

    ReportMan setColumns(List<String> columnList){
        columnList.each{ String columnString ->
            system.option.sheetInfoMap.columnOptMap[columnString] = new ColumnOption(index: ReportUtil.getColumnIndex(columnString))
        }
        return this
    }

    ReportMan setExcludeColumn(String... excludeColumnFieldNames){
        system.option.excludeColumnFieldNameList = excludeColumnFieldNames.toList()
        return this
    }

    ReportMan addAdditionalColumnOption(Map<String, ColumnOption> additionalColumnOptionMap){
        system.option.additionalColumnOptionMap = additionalColumnOptionMap
        return this
    }

    ReportMan setModeOnlyHeader(Boolean modeOnlyHeader){
        system.option.modeOnlyHeader = modeOnlyHeader
        return this
    }



    /**************************************************
     *
     * Read
     *  - Excel File  ==> Data (List)
     *
     **************************************************/
    List<Map<String, ?>> toRowList(){
        File file = new File(system.option.fileName)
        InputStream inputStream = new FileInputStream(file)
        return toRowList(inputStream)
    }

    List<?> toRowList(def instance){
        File file = new File(system.option.fileName)
        InputStream inputStream = new FileInputStream(file)
        toRowList(inputStream, instance)
    }

    List<?> toRowList(InputStream inputStream){
        return toRowList(inputStream, [:])
    }

    List<?> toRowList(InputStream inputStream, def instance){
        List allRowList = []
        excelToData(inputStream, instance){ String sheetName, List rowList ->
            allRowList += rowList
        }
        return allRowList
    }



    /**************************************************
     *
     * Read
     *  - Excel File  ==> Data (Map)
     *
     **************************************************/
    Map<String, List<?>> toSheetMap(){
        File file = new File(system.option.fileName)
        InputStream inputStream = new FileInputStream(file)
        return toSheetMap(inputStream)
    }

    Map<String, List<?>> toSheetMap(def instance){
        File file = new File(system.option.fileName)
        InputStream inputStream = new FileInputStream(file)
        return toSheetMap(inputStream, instance)
    }

    Map<String, List<?>> toSheetMap(InputStream inputStream){
        return toSheetMap(inputStream, null)
    }

    Map<String, List<?>> toSheetMap(InputStream inputStream, def instance){
        Map sheetMap = [:]
        excelToData(inputStream, instance){ String sheetName, List rowList ->
            sheetMap[sheetName] = rowList
        }
        return sheetMap
    }



    private boolean excelToData(InputStream inputStream, def instance, Closure closure){
        // 1. Get excel file
        Workbook workbook = WorkbookFactory.create( inputStream )
        int sheetSize = workbook.numberOfSheets

        // 2. Get sheet.
        ReportSheetInfo sheetInfo
        Sheet sheet
        String sheetName
        List<?> rows

        for (int sheetIndex=0; sheetIndex<sheetSize; sheetIndex++){
            sheet = workbook.getSheetAt(sheetIndex)
            sheetName = sheet.getSheetName()
            //Generate SheetOption, ColumnOptions
            sheetInfo = makeSheetInfo(sheetName, system.option, instance)
            rows = extractOneSheet(sheet, sheetInfo, sheetName, instance, closure)
        }
        return rows
    }

    private List extractOneSheet(Sheet sheet, ReportSheetInfo sheetInfo, String sheetName, def instance, Closure closure){
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        Row row
        List rows = []
        int rowIndex = 0
        for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext();){
            // 4. Get Cells.
            row = rowsIT.next()
            if (rowIndex++ < csi.dataStartY){
                continue
            }
            rows << extractOneRow(row, sheetInfo, sheetName, instance)
        }

        closure(sheetName, rows)
        return rows
    }

    private def extractOneRow(Row row, ReportSheetInfo sheetInfo, String sheetName, def instance){
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo
        Object mapperInstance = sheetInfo.mapperInstance
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap

        def rowDto = (instance != null) ? instance.getClass().newInstance() : (mapperInstance != null) ? mapperInstance.getClass().newInstance() : [:]

        columnOptMap.each{ String attr, ColumnOption cOpt ->
            // 5. Get Only Mapping Cell
            Cell cell = (cOpt.index > -1) ? row.getCell(csi.dataStartX + cOpt.index) : null
            if (cell){
                if (attr){
                    Class clazz = rowDto?.metaClass?.properties?.find{ it.name == attr }?.type
                    cell.setCellType(Cell.CELL_TYPE_STRING)
                    if (clazz){
                        if (clazz == Integer.class){
                            String value = cell.getStringCellValue()
                            rowDto[attr] = value.isNumber() ? ((Double)Double.parseDouble(value)).intValue() : null
                        }else{
                            rowDto[attr] = cell.getStringCellValue()
                        }
                    }else{
                        rowDto[attr] = cell.getStringCellValue()
                    }
                }
            }else if (cOpt.isSheetNameField){
                rowDto[attr] = sheetName
            }
        }
        return rowDto
    }



    /**************************************************
     * Write
     *  - Excel File  <== Data
     *  - Download Example>
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
        String sheetName
        List rowList
        allRowList.each{
            sheetName = it[sheetFieldName]
            rowList = sheetMap[sheetName]
            if (!rowList)
                sheetMap[it[sheetFieldName]] = [it]
            else
                rowList << it
        }
        return write(sheetMap)
    }

    boolean write(Map sheetMap){
        //Make SheetNames (null => No Sheet Name)
        sheetMap.findAll{ String sheetName, List dataRowList ->
            return !sheetName
        }.each{ String sheetName, List dataRowList ->
            if (!sheetMap['No Sheet Name'])
                sheetMap['No Sheet Name'] = []
            (sheetMap['No Sheet Name'] as List).addAll(dataRowList)
        }
        sheetMap.remove(null)

        //Create WorkBook
        system.workbook = createWorkBook(sheetMap, system.option)

        //Write Excel File
        return write(system.workbook, system.getOutputStream())
    }

    boolean write(XSSFWorkbook workbook){
        return write(workbook, system.getOutputStream())
    }

    boolean write(OutputStream outputStream){
        return write(system.workbook, outputStream)
    }

    boolean write(XSSFWorkbook workbook, OutputStream outputStream){
        workbook.write(outputStream)
        outputStream.flush()
        outputStream.close()
        return true
    }



    private XSSFWorkbook createWorkBook(Map sheetMap, ReportOption option){
        ReportSheetInfo sheetInfo

        XSSFWorkbook workbook = new XSSFWorkbook();
        sheetMap.each{ String sheetName, List dataRowList ->
            def instance
            if (dataRowList != null && dataRowList.size() > 0)
                instance = dataRowList.get(0)
            //Generate SheetOption, ColumnOptions
            sheetInfo = makeSheetInfo(sheetName, option, instance)
            //Create Sheet
            sheetInfo.sheet = createSheet(workbook, sheetName, dataRowList, option, sheetInfo)
        }
        return workbook
    }

    ReportSheetInfo makeSheetInfo(String sheetName, ReportOption option){
        return makeSheetInfo(sheetName, option)
    }

    ReportSheetInfo makeSheetInfo(String sheetName, ReportOption option, def instance){
        //Force
        SheetOption sheetOpt = option?.defaultSheetInfo?.sheetOpt
        Map<String, ColumnOption> columnOptMap = option?.defaultSheetInfo?.columnOptMap
        Object mapperInstance = option?.defaultSheetInfo?.mapperInstance

        //By SheetName
        if (!option)
            option = new ReportOption()
        if (!option.sheetInfoMap)
            option.sheetInfoMap = [:]

        ReportSheetInfo reportSheetInfo = option.sheetInfoMap[sheetName]
        if (!reportSheetInfo){
            option.sheetInfoMap[sheetName] = reportSheetInfo = new ReportSheetInfo()
            if (instance){
                //From instance's annotation
                reportSheetInfo.sheetOpt = sheetOpt ?: AnnotationExtractor.generateSheetOption(instance)
                reportSheetInfo.columnOptMap = columnOptMap ?: AnnotationExtractor.generateColumnOptionMap(instance)
                reportSheetInfo.mapperInstance = instance ?: mapperInstance
            }else{
                //From code
                reportSheetInfo.sheetOpt = sheetOpt ?: AnnotationExtractor.generateSheetOption(mapperInstance)
                reportSheetInfo.columnOptMap = columnOptMap ?: AnnotationExtractor.generateSheetOption(mapperInstance)
                reportSheetInfo.mapperInstance = mapperInstance
            }
        }

        if (!reportSheetInfo.calculatedSheetInfo)
            reportSheetInfo.calculatedSheetInfo = ReportUtil.calculatePositionData(reportSheetInfo)

        return reportSheetInfo
    }

    private XSSFSheet createSheet(XSSFWorkbook workbook, String sheetName, List dataRowList, ReportOption option, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        //Create Sheet
        XSSFSheet sheet = workbook.createSheet(sheetName)

        //Calculate
        ReportUtil.calculateStyle(workbook, sheetInfo)
        csi.dataSize = dataRowList.size()
        csi.replaceMap = ReportUtil.generateReplaceMap(sheetInfo)

        //SIZE(Sheet) - "왜 width는 안먹냐"
        sheetStyleHandler.setDefaultSize(sheet, sheetInfo)

        /** Row - Header **/
        if (csi.modeHeader)
            writeHeader(sheet, sheetInfo)
        if (option.modeOnlyHeader)
            return

        /** Row - Data **/
        writeData(sheet, dataRowList, sheetInfo)

        /** Additional Style **/
        sheetStyleHandler.autoFilter(sheet, sheetInfo)
        sheetStyleHandler.freezePane(sheet, sheetInfo)
        sheetStyleHandler.sheetHighlightStyle(sheet, sheetInfo)
        columnStyleHandler.columnHighlightStyle(sheet, sheetInfo)
        return sheet
    }



    private XSSFRow writeHeader(Sheet sheet, ReportSheetInfo sheetInfo){
        //Create Row(Header)
        return writeOneHeaderRow(sheet, sheetInfo)
    }

    private XSSFRow writeOneHeaderRow(Sheet sheet, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        int headerStartX = csi.headerStartX
        int headerStartY = csi.headerStartY
        Map<String, CellStyle> cellStyleMap = sheetOpt.columnHeaderCellStyleMap
        XSSFRow headerRow = sheet.createRow(headerStartY)
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            int index = cOpt.index
            String headerName = cOpt.headerName
            if (headerName && index != -1){
                XSSFCell cell = headerRow.createCell(headerStartX + index)
                rowWriteHandler.setHeaderData(cell, headerName)
                rowWriteHandler.setHeaderStyle(cell, headerName, cellStyleMap[attr])
                //SIZE - WIDTH
                sheetStyleHandler.setColumnWidth(sheet, csi.headerStartX + index, cOpt.width)
            }
        }
        //SIZE - HEIGHT
        sheetStyleHandler.setRowHeight(headerRow, sheetOpt.headerHeight)
        return headerRow
    }

    private void writeData(Sheet sheet, List dataRowList, ReportSheetInfo sheetInfo){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        //Option - DataTwoToneStyle
        boolean modeTwoToneChange = (!!sheetOpt.dataTwoToneStyle)
        boolean modeTwoToneToggleChange
        List<String> pkList
        Map pkMap = [:]
        if (modeTwoToneChange){
            pkList = sheetOpt.dataTwoToneStyle.option.pk.split("\\s*,\\s*")
            if (pkList[0])
                pkList.each{ pkMap[it] = null }
            else
                modeTwoToneToggleChange = true
        }

        //CellStyleMap
        Map<String, CellStyle> cellStyleMap = (modeTwoToneChange) ? sheetOpt.columnDataTwoToneCellStyleMap : sheetOpt.columnDataCellStyleMap

        //Row(Data)
        dataRowList.eachWithIndex{ rowData, rowIndex ->
            //Check DataTwoToneChange
            if (modeTwoToneChange){
                boolean isTwoToneChange = rowWriteHandler.checkTwoTone(columnOptMap, pkMap, rowData)
                if (isTwoToneChange || modeTwoToneToggleChange)
                    cellStyleMap = (cellStyleMap == sheetOpt.columnDataCellStyleMap) ? sheetOpt.columnDataTwoToneCellStyleMap : sheetOpt.columnDataCellStyleMap
            }
            writeOneDataRow(sheet, rowData, rowIndex, sheetInfo, cellStyleMap)
        }
    }

    private XSSFRow writeOneDataRow(Sheet sheet, Object rowData, int rowIndex, ReportSheetInfo sheetInfo, Map<String, CellStyle> cellStyleMap){
        SheetOption sheetOpt = sheetInfo.sheetOpt
        Map<String, ColumnOption> columnOptMap = sheetInfo.columnOptMap
        CalculatedSheetInfo csi = sheetInfo.calculatedSheetInfo

        XSSFRow row = sheet.createRow(rowIndex + csi.dataStartY)
        columnOptMap.each{ String attr, ColumnOption cOpt ->
            int index = cOpt.index
            if (index != -1){
                XSSFCell cell = row.createCell(index + csi.dataStartX);
                rowWriteHandler.setData(cell, rowData[attr])
                rowWriteHandler.setStyle(cell, rowData[attr], cellStyleMap[attr])
            }
        }
        return row
    }


}
