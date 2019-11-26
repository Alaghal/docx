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
        //SpringApplication.run(DocxApplication.class, args);
        testWorkTabl workTabl = new testWorkTabl();
        WordprocessingMLPackage template =  workTabl.getTemplate("C:\\Users\\Серяковы\\IdeaProjects\\docx\\src\\main\\resources\\templates\\template.docx");


        List<SectionData> sectionDataList = new ArrayList<>();

        Map<String,String> repl1 = new HashMap();
        repl1.put("SJ_FUNCTION", "function1");
        repl1.put("SJ_DESC", "desc1");
        repl1.put("SJ_PERIOD", "period1");

        Map<String,String> repl2 = new HashMap();
        repl2.put("SJ_FUNCTION", "function2");
        repl2.put("SJ_DESC", "desc2");
        repl2.put("SJ_PERIOD", "period2");

        Map<String,String> repl3 = new HashMap();
        repl3.put("SJ_FUNCTION", "function3");
        repl3.put("SJ_DESC", "desc3");
        repl3.put("SJ_PERIOD", "period3");
        
        Map<String,String> dataToAddForParagraps = new HashMap();
        dataToAddForParagraps.put("<VALUE>","value");
        dataToAddForParagraps.put("<VALUE2>","value1");
    
        Map<String,String> dataToAddForParagraps2 = new HashMap();
        dataToAddForParagraps2.put("<VALUE>","value");
        dataToAddForParagraps2.put("<VALUE2>","value2");
        
        Map<String,String> dataToAddForParagraps3 = new HashMap();
        dataToAddForParagraps3.put("<VALUE>","value");
        dataToAddForParagraps3.put("<VALUE2>","value3");

        sectionDataList.add(new SectionData(Arrays.asList(repl1,repl2,repl3),dataToAddForParagraps));
        sectionDataList.add(new SectionData(Arrays.asList(repl1,repl2,repl3),dataToAddForParagraps2));
        sectionDataList.add(new SectionData(Arrays.asList(repl1,repl2,repl3),dataToAddForParagraps3));

        workTabl.fillTemplateBySanctionsData(template,sectionDataList,"START_SECTIONS", "END_SECTION", "<",">");
          List<Object> objectList = template.getMainDocumentPart().getContent();
        workTabl.writeDocxToStream(template,"template2.docx");
        //workTabl.replaceTable(new String[]{"SJ_FUNCTION","SJ_DESC","SJ_PERIOD"}, Arrays.asList(repl1,repl2,repl3),template);


    }

}
