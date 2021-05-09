package jaemisseo.man.reportman.model

class ColumnOption extends StyleOption {

    int index
    //SIZE
    int width = 5000
    int height
    //
    String headerName
    boolean isSheetNameField
    //
    StyleOption headerStyle
    StyleOption dataStyle
    StyleOption highlightStyle

}
