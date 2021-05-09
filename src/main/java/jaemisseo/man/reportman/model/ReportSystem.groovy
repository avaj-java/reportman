package jaemisseo.man.reportman.model


import org.apache.poi.xssf.usermodel.XSSFWorkbook

class ReportSystem {

    XSSFWorkbook workbook

//    File file
//    String fileName
//    OutputStream outputStream

    ReportOption option


    OutputStream getOutputStream(){
        OutputStream os
        if (option.outputStream) {
            os = option.outputStream
        }else{
            if (option.file){
                os = new FileOutputStream( option.file )
            }else if (option.fileName){
                os = new FileOutputStream( new File(option.fileName) )
            }else{
                os = null
            }
        }
        return os
    }


}
