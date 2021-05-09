package jaemisseo.man.reportman.model

import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.ReportSheetInfo

/**
 * User can use
 */
class ReportOption {

    File file
    String fileName
    OutputStream outputStream
    InputStream inputStream

    Map<String, ColumnOption> additionalColumnOptionMap
    List<String> excludeColumnFieldNameList
    Boolean modeOnlyHeader = false

    ReportSheetInfo defaultSheetInfo
    Map<String, ReportSheetInfo> sheetInfoMap

}
