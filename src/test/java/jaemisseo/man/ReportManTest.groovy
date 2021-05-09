package jaemisseo.man

import jaemisseo.man.bean.ReportManTestBean
import jaemisseo.man.bean.ReportManTestBean2
import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.ReportOption
import jaemisseo.man.reportman.model.ReportSheetInfo
import jaemisseo.man.reportman.model.SheetOption
import jaemisseo.man.reportman.model.StyleOption
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.util.HSSFColor
import org.junit.After
import org.junit.Before
import org.junit.Test

class ReportManTest {

    @Before
    void init(){
    }

    @After
    void after(){
    }

    String testPath = 'build/tmp/'

    /*************************
     * Data
     *************************/
    List<ReportManTestBean> emptyData = []

    List<ReportManTestBean> data = [
            new ReportManTestBean(objectId:1, groupName:"Group1", attributeName:"ATT1_1", description:"Hello, nice to meet you. and you?"),
            new ReportManTestBean(objectId:2, groupName:"Group1", attributeName:"ATT1_2", description:"May i help you?"),
            new ReportManTestBean(objectId:3, groupName:"Group2", attributeName:"ATT2_1", description:"May i have a attention please."),
            new ReportManTestBean(objectId:3, groupName:"Group3", attributeName:"ATT3_1", description:"can i love you?"),
    ]




    @Test
    void emptyList_to_excel__no_style(){
        ReportOption option = new ReportOption(
                fileName: "$testPath/emptyList_to_excel__no_style.xlsx",
                defaultSheetInfo: new ReportSheetInfo(
                        sheetOpt: new SheetOption(),
                        columnOptMap: [
                                attributeName: new ColumnOption(index:0, headerName: "SOME1"),
                                description: new ColumnOption(index:1, headerName: "SOME222")
                        ],
                )
        )
        new ReportMan(option).write(emptyData)
        assert new File(option.fileName).exists()
    }

    @Test
    void emptyList_to_excel__stylish(){
        ReportOption option = new ReportOption(
                fileName: "$testPath/emptyList_to_excel__stylish.xlsx",
                defaultSheetInfo: new ReportSheetInfo(
                        sheetOpt: new SheetOption(),
                        columnOptMap: [
                                attributeName: new ColumnOption(
                                        index:0,
                                        headerName: "Attribute Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                ),
                                description: new ColumnOption(
                                        index:1,
                                        headerName: "Description Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                )
                        ],
                )
        )
        new ReportMan(option).write(emptyData)
        assert new File(option.fileName).exists()
    }



    @Test
    void list_to_excel_to_list(){
        ReportOption option = new ReportOption(
                fileName: "$testPath/list_to_excel.xlsx",
                defaultSheetInfo: new ReportSheetInfo(
                        sheetOpt: new SheetOption(),
                        columnOptMap: [
                                attributeName: new ColumnOption(
                                        index:0,
                                        headerName: "Attribute Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                ),
                                description: new ColumnOption(
                                        index:1,
                                        headerName: "Description Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                )
                        ],
                )
        )

        /** Write as 1 Sheet **/
        new ReportMan(option).write(data)
        assert new File(option.fileName).exists()

        /** Read as Map **/
        //List<Map<String, ?>>
        List<Map> rowList = new ReportMan(option).toRowList()
        assert rowList instanceof List
        assert rowList.size() == 4
        assert rowList.every{ it instanceof Map }

        //Map<String, Map<String, ?>>
        Map<String, List<ReportManTestBean>> sheetMap = new ReportMan(option).toSheetMap()
        assert sheetMap instanceof Map
        assert sheetMap.size() == 1
        assert sheetMap.every{ sheetName, rows -> sheetName instanceof String && rows instanceof List }

        /** Read as Some Instance **/
        //List<?>
        rowList = new ReportMan(option).toRowList(new ReportManTestBean())
        assert rowList instanceof List
        assert rowList.size() == 4
        assert rowList.every{ it instanceof ReportManTestBean }

        //Map<String, ?>
        sheetMap = new ReportMan(option).toSheetMap(new ReportManTestBean())
        assert sheetMap instanceof Map
        assert sheetMap.size() == 1
        assert sheetMap.every{ sheetName, rows -> sheetName instanceof String && rows instanceof List }
    }

    @Test
    void list_to_excelAsMultiSheet_to_list(){
        ReportOption option = new ReportOption(
                fileName: "$testPath/list_to_excelAsMultiSheet.xlsx",
                defaultSheetInfo: new ReportSheetInfo(
                        sheetOpt: new SheetOption(),
                        columnOptMap: [
                                attributeName: new ColumnOption(
                                        index:0,
                                        headerName: "Attribute Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                ),
                                description: new ColumnOption(
                                        index:1,
                                        headerName: "Description Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                )
                        ],
                )
        )

        /** Write as Multi Sheet **/
        new ReportMan(option).write(data, "groupName")
        assert new File(option.fileName).exists()

        /** Read as Map **/
        //List<Map<String, ?>>
        List<Map> rowList = new ReportMan(option).toRowList()
        assert rowList instanceof List
        assert rowList.size() == 4
        assert rowList.every{ it instanceof Map }

        //Map<String, Map<String, ?>>
        Map<String, List<ReportManTestBean>> sheetMap = new ReportMan(option).toSheetMap()
        assert sheetMap instanceof Map
        assert sheetMap.size() == 3
        assert sheetMap.every{ sheetName, rows -> sheetName instanceof String && rows instanceof List }

        /** Read as Some Instance **/
        //List<?>
        rowList = new ReportMan(option).toRowList(new ReportManTestBean())
        assert rowList instanceof List
        assert rowList.size() == 4
        assert rowList.every{ it instanceof ReportManTestBean }

        //Map<String, ?>
        sheetMap = new ReportMan(option).toSheetMap(new ReportManTestBean())
        assert sheetMap instanceof Map
        assert sheetMap.size() == 3
        assert sheetMap.every{ sheetName, rows -> sheetName instanceof String && rows instanceof List }
    }

    @Test
    void list_to_excelAsMultiSheet_to_list_with_multi_class(){
        ReportOption option = new ReportOption(
                fileName: "$testPath/list_to_excelAsMultiSheet_to_list_with_multi_class.xlsx",
                defaultSheetInfo: new ReportSheetInfo(
                        mapperInstance: new ReportManTestBean(),
                        sheetOpt: new SheetOption(),
                        columnOptMap: [
                                attributeName: new ColumnOption(
                                        index:0,
                                        headerName: "Attribute Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                ),
                                description: new ColumnOption(
                                        index:1,
                                        headerName: "Description Something",
                                        headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_YELLOW.index, bold:true, alignment:HSSFCellStyle.ALIGN_CENTER, verticalAlignment:HSSFCellStyle.VERTICAL_CENTER),
                                )
                        ],
                ),
                sheetInfoMap: [
                        "Group1": new ReportSheetInfo(
                                mapperInstance: new ReportManTestBean2(),
                                sheetOpt: new SheetOption(),
                                columnOptMap: [
                                        attributeName: new ColumnOption(
                                                index:0,
                                                headerName: "Attribute Something",
                                                headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_GREEN.index),
                                        ),
                                        description: new ColumnOption(
                                                index:1,
                                                headerName: "Description Something",
                                                headerStyle: new StyleOption(foreground:HSSFColor.LIGHT_GREEN.index),
                                        )
                                ],
                        ),
                ]
        )

        /** Write as Multi Sheet **/
        new ReportMan(option).write(data, "groupName")
        assert new File(option.fileName).exists()

        /** Read as Map **/
        //List<Map<String, ?>>
        List<Map> rowList = new ReportMan(option).toRowList()
        assert rowList instanceof List
        assert rowList.size() == 4
        assert rowList.every{ it instanceof Map }

        //Map<String, Map<String, ?>>
        Map<String, List<ReportManTestBean>> sheetMap = new ReportMan(option).toSheetMap()
        assert sheetMap instanceof Map
        assert sheetMap.size() == 3
        assert sheetMap.every{ sheetName, rows -> sheetName instanceof String && rows instanceof List }

        /** Read as Some Instance **/
        //List<?>
        //None

        //Map<String, ?>
        sheetMap = new ReportMan(option).toSheetMap()
        assert sheetMap instanceof Map
        assert sheetMap.size() == 3
        assert sheetMap.every{ sheetName, rows -> sheetName instanceof String && rows instanceof List }
        assert sheetMap.every{ sheetName, rows ->
            Object mapperInstance = option.sheetInfoMap[sheetName].mapperInstance
            rows.every{ it.getClass().isAssignableFrom(mapperInstance.getClass()) }
        }
    }

}
