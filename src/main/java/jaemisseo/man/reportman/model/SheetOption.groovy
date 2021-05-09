package jaemisseo.man.reportman.model

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.usermodel.XSSFCellStyle

class SheetOption extends StyleOption {

    //SIZE
    int width = -1
    int height = 230
    int headerHeight = 500
    //Position
    String headerPosition
    String dataPosition

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
