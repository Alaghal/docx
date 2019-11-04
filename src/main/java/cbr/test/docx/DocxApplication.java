package cbr.test.docx;

import cbr.test.docx.services.Editors.docx.SectionData;
import cbr.test.docx.services.Editors.docx.testWorkTabl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.xml.bind.JAXBException;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

@SpringBootApplication
public class DocxApplication {

    public static void main(String[] args) throws IOException, Docx4JException, JAXBException {
        SpringApplication.run(DocxApplication.class, args);
        testWorkTabl workTabl = new testWorkTabl();
        WordprocessingMLPackage template =  workTabl.getTemplate("C:\\Users\\Серяковы\\Desktop\\template.docx");


        List<SectionData> sectionDataList = new ArrayList<>();

        Map<String,String> repl1 = new HashMap<String, String>();
        repl1.put("SJ_FUNCTION", "function1");
        repl1.put("SJ_DESC", "desc1");
        repl1.put("SJ_PERIOD", "period1");

        Map<String,String> repl2 = new HashMap<String,String>();
        repl2.put("SJ_FUNCTION", "function2");
        repl2.put("SJ_DESC", "desc2");
        repl2.put("SJ_PERIOD", "period2");

        Map<String,String> repl3 = new HashMap<String,String>();
        repl3.put("SJ_FUNCTION", "function3");
        repl3.put("SJ_DESC", "desc3");
        repl3.put("SJ_PERIOD", "period3");

        sectionDataList.add(new SectionData(Arrays.asList(repl1,repl2,repl3),null));
        sectionDataList.add(new SectionData(Arrays.asList(repl1,repl2,repl3),null));
        sectionDataList.add(new SectionData(Arrays.asList(repl1,repl2,repl3),null));

        workTabl.fillTemplateBySanctionsData(template,sectionDataList,"<StartSection>", "<EndSection>");

        workTabl.writeDocxToStream(template,"C:\\Users\\Серяковы\\Desktop\\template2.docx");
        //workTabl.replaceTable(new String[]{"SJ_FUNCTION","SJ_DESC","SJ_PERIOD"}, Arrays.asList(repl1,repl2,repl3),template);


    }

}
