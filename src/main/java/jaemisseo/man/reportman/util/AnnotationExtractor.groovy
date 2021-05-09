package jaemisseo.man.reportman.util

import jaemisseo.man.reportman.model.ReportOption
import jaemisseo.man.annotation.ReportColumn
import jaemisseo.man.annotation.ReportColumnDataStyle
import jaemisseo.man.annotation.ReportColumnHeaderStyle
import jaemisseo.man.annotation.ReportColumnHighlightStyle
import jaemisseo.man.annotation.ReportSheet
import jaemisseo.man.annotation.ReportSheetDataStyle
import jaemisseo.man.annotation.ReportSheetDataTwoToneStyle
import jaemisseo.man.annotation.ReportSheetHeaderStyle
import jaemisseo.man.annotation.ReportSheetHighlightStyle
import jaemisseo.man.annotation.ReportSheetName
import jaemisseo.man.annotation.ReportSheetStyle
import jaemisseo.man.reportman.model.ColumnOption
import jaemisseo.man.reportman.model.SheetOption
import jaemisseo.man.reportman.model.StyleOption

import java.lang.annotation.Annotation
import java.lang.reflect.Field

class AnnotationExtractor {


    static SheetOption generateSheetOption(def instance){
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

    static Map generateColumnOptionMap(def instance){
        generateColumnOptionMap(instance, null)
    }

    static Map generateColumnOptionMap(def instance, ReportOption option){
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
                if (!option?.excludeColumnFieldNameList?.contains(fieldName)){
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
        option?.additionalColumnOptionMap?.each{ String fieldName, ColumnOption columnOption ->
            columnOptMap[fieldName] = columnOption
        }
        //Sorting
        columnOptMap.sort{ a, b ->
            a.value.index <=> b.value.index
        }
        return columnOptMap
    }

    static StyleOption generateStyleOption(Annotation ant){
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

    static StyleOption generateHightlightStyleOption(Annotation ant){
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

}
