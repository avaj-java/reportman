package jaemisseo.man.bean

import jaemisseo.man.ReportMan
import jaemisseo.man.annotation.*
import org.apache.poi.hssf.usermodel.HSSFCellStyle
import org.apache.poi.hssf.util.HSSFColor

/**
 * Created by sujkim on 2017-02-02.
 */
@ReportSheet(height=230, headerHeight=500, freezePane=ReportMan.RANGE_AUTO, autoFilter=ReportMan.RANGE_AUTO)
@ReportSheetStyle(fontSize=9, border=HSSFCellStyle.BORDER_THIN)
@ReportSheetHeaderStyle(foreground=HSSFColor.LIGHT_YELLOW.index, bold=true, alignment=HSSFCellStyle.ALIGN_CENTER, verticalAlignment=HSSFCellStyle.VERTICAL_CENTER)
@ReportSheetDataStyle(foreground=HSSFColor.LIGHT_GREEN.index)
@ReportSheetDataTwoToneStyle(pk="className")
//@ReportSheetHighlightStyle(condition='$${classId}${dataStartRow}=""', range=ReportMan.RANGE_AUTO, color=HSSFColor.DARK_RED.index, border=HSSFCellStyle.BORDER_THICK, background=HSSFColor.PINK.index)
class ManTestBean {

    Integer objectId

    @ReportSheetName
    String groupName

    @ReportColumnHighlightStyle(condition='$${0}${dataStartRow}=""', range=ReportMan.RANGE_DATA_ALL, background=HSSFColor.DARK_RED.index)
    @ReportColumn(index=0, headerName="클래스명", width=9000)
    String className


    @ReportColumn(index=2, headerName="어트리뷰트명", width=7000)
    String attributeName

    @ReportColumn(index=7, headerName="설명", width=12000)
    String description

    @ReportColumn(index=6, headerName="테이블명", width=5500)
    String tableName

    @ReportColumn(index=1, headerName="CLASS_ID", width=2000)
    Integer classId

    @ReportColumn(index=3, headerName="PROPERTY_ID", width=2000)
    Integer propertyId

    @ReportColumn(index=4, headerName="컬럼명", width=5000)
    String columnName

    Date createDt

    String cusr

    Integer page

    Integer pageSize

    void setTableName(String tableName){ this.tableName = tableName?.toUpperCase() }
    void setColumnName(String columnName){ this.columnName = columnName?.toUpperCase() }
}
