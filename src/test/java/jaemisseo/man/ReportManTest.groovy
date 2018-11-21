package jaemisseo.man

import jaemisseo.man.bean.ManTestBean
import org.junit.After
import org.junit.Before
import org.junit.Ignore
import org.junit.Test

/**
 * Created by sujkim on 2017-02-08.
 */
class ReportManTest {

    @Before
    void init(){
    }

    @After
    void after(){
    }

    Map dbInfo(){ return [ip:'127.0.0.1', db:'orcl', uaser:'tester', password:'tester'] }



    @Test
    @Ignore
    void "DB To Excel"(){
        //DB to SheetMap
        List resultList = []
        //SheetMap To Excel
        new ReportMan('meta3mapping.xlsx').write(resultList, 'groupName')
    }

    @Test
    @Ignore
    void "Excel To DB"(){
        String createUser = "Tester/tester"
        InputStream inputStream = new FileInputStream(new File("meta3mapping.xlsx"))
        //Excel To RowList
        List<ManTestBean> allRowList = new ReportMan().toRowList(inputStream, new ManTestBean())
    }

}
