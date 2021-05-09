package jaemisseo.man

import jaemisseo.man.bean.ReportManTestBean
import org.junit.After
import org.junit.Before
import org.junit.Test

class ReportManTest_by_annotation {

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
    void emptyList_to_excel(){
        String filePath = "$testPath/emptyList_to_excel_annotationOption.xlsx"

        List<ReportManTestBean> datas = [new ReportManTestBean()]

        new ReportMan(filePath).write(datas)
        assert new File(filePath).exists()
    }

    @Test
    void list_to_excel_to_list(){
        String filePath = "$testPath/list_to_excel_annotationOption.xlsx"

        new ReportMan(filePath).write(data, 'groupName')
        assert new File(filePath).exists()

        //List
        InputStream inputStream = new FileInputStream( new File(filePath) )
        List<ReportManTestBean> allRowList = new ReportMan().toRowList(inputStream, new ReportManTestBean())
        assert allRowList instanceof List
        assert allRowList.size() == 4

        //Map
        inputStream = new FileInputStream( new File(filePath) )
        Map<String, List<ReportManTestBean>> map = new ReportMan().toSheetMap(inputStream, new ReportManTestBean())
        assert map instanceof Map
        assert map.size() == 3
    }



}
