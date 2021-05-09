package jaemisseo.man.reportman.model

import org.apache.poi.ss.usermodel.FillPatternType

class StyleOption {

    //POIs STYLE
    short color = -1
    short fontHeightInPoints = -1
    boolean bold = false
    boolean italic = false
    boolean wrapText = false
    short alignment = -1
    short verticalAlignment = -1
    short fillForegroundColor = -1
    short fillBackgroundColor = -1
    FillPatternType fillPattern = FillPatternType.NO_FILL
    short borderTop = -1
    short borderBottom = -1
    short borderLeft = -1
    short borderRight = -1
    //ReportMan CUSTOM PROEPRTIES
    short fontSize = -1
    short border = -1
    short foreground = -1
    short background = -1
    //OPTION
    Map option = [:]

}
