package cbr.test.docx;

import cbr.test.docx.services.Editors.docx.component.SectionData;
import cbr.test.docx.services.Editors.docx.service.GenerateSectionsWord;
import cbr.test.docx.services.Editors.docx.service.impl.GenerateSectionsWordImpl;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.xml.bind.JAXBException;
import java.io.IOException;
import java.util.*;

@SpringBootApplication
public class DocxApplication {

    public static void main(String[] args) throws IOException, Docx4JException, JAXBException {
        //SpringApplication.run(DocxApplication.class, args);



        List<SectionData> sectionDataList = new ArrayList<>();

        Map<String,Object> repl1 = new HashMap();
        repl1.put("P1_PCI", "11");
        repl1.put("P1_AI", "12");
        repl1.put("P1_AIO", "13");
        repl1.put("P2_PCI", "14");
        repl1.put("P2_AI", "15");
        repl1.put("P2_AIO", "16");
        repl1.put("P3_PCI", "17");
        repl1.put("P3_AI", "18");
        repl1.put("P3_AIO", "19");
        repl1.put("P4_PCI", "110");
        repl1.put("P4_AI", "111");
        repl1.put("P4_AIO", "112");
        repl1.put("TOT_PCI", "113");
        repl1.put("TOT_AI", "114");
        repl1.put("TOT_AIO", "115");
        repl1.put("P1_PO", "116");
        repl1.put("P1_TA", "117");
        repl1.put("P1_TB", "118");
        repl1.put("P2_PO", "119");
        repl1.put("P2_TA", "120");
        repl1.put("P2_TB", "121");
        repl1.put("P3_PO", "122");
        repl1.put("P3_TA", "123");
        repl1.put("P3_TB", "124");
        repl1.put("P4_PO", "125");
        repl1.put("P4_TA", "126");
        repl1.put("P4_TB", "127");
        repl1.put("TOT_PO", "128");
        repl1.put("TOT_TA", "129");
        repl1.put("TOT_TB", "130");
        repl1.put("LASTQUARTERNAME", "131");
        repl1.put("CURRENTQUARTERNAME", "132");
        repl1.put("CLAIMSINCLUDING_LQBEP", "133");
        repl1.put("CLAIMSINCLUDING_CQBEP", "134");
        repl1.put("CLAIMSCERTIFIEDBYMORTGAGES_LQBEP", "133");
        repl1.put("CLAIMSCERTIFIEDBYMORTGAGES_CQBEP", "134");
        repl1.put("CLAIMSCONSTRUCTIONNOTCOMPLETED_LQBEP", "135");
        repl1.put("CLAIMSCONSTRUCTIONNOTCOMPLETED_CQBEP", "136");
        repl1.put("CLAIMSBYRESIDENTIALMORTGAGES_LQBEP", "137");
        repl1.put("CLAIMSBYRESIDENTIALMORTGAGES_CQBEP", "138");
        repl1.put("CLAIMSRESIDENTIALPREMISES_LQBEP", "139");
        repl1.put("CLAIMSRESIDENTIALPREMISES_CQBEP", "140");
        repl1.put("PARTICIPATIONCERTIFICATES_LQBEP", "141");
        repl1.put("PARTICIPATIONCERTIFICATES_CQBEP", "142");
        repl1.put("CASHINCLUDING_LQBEP", "143");
        repl1.put("CASHINCLUDING_CQBEP", "144");
        repl1.put("CURRENCYOFRF_LQBEP", "145");
        repl1.put("CURRENCYOFRF_CQBEP", "146");
        repl1.put("FOREIGNCURRENCY_LQBEP", "147");
        repl1.put("FOREIGNCURRENCY_CQBEP", "148");
        repl1.put("SIZEMORTGAGECOVERAGETOTAL_LQBEP", "149");
        repl1.put("SIZEMORTGAGECOVERAGETOTAL_CQBEP", "150");
        repl1.put("ESTIMATEDCCOSTPERISU_LQBEP", "151");
        repl1.put("ESTIMATEDCOSTPERISU_CQBEP", "152");

        Map<String,Object> repl2 = new HashMap();
        repl2.put("P1_PCI", "21");
        repl2.put("P1_AI", "22");
        repl2.put("P1_AIO", "23");
        repl2.put("P2_PCI", "24");
        repl2.put("P2_AI", "25");
        repl2.put("P2_AIO", "26");
        repl2.put("P3_PCI", "27");
        repl2.put("P3_AI", "28");
        repl2.put("P3_AIO", "29");
        repl2.put("P4_PCI", "210");
        repl2.put("P4_AI", "211");
        repl2.put("P4_AIO", "212");
        repl2.put("TOT_PCI", "213");
        repl2.put("TOT_AI", "214");
        repl2.put("TOT_AIO", "215");
        repl2.put("P1_PO", "216");
        repl2.put("P1_TA", "217");
        repl2.put("P1_TB", "218");
        repl2.put("P2_PO", "219");
        repl2.put("P2_TA", "220");
        repl2.put("P2_TB", "221");
        repl2.put("P3_PO", "222");
        repl2.put("P3_TA", "223");
        repl2.put("P3_TB", "224");
        repl2.put("P4_PO", "225");
        repl2.put("P4_TA", "226");
        repl2.put("P4_TB", "227");
        repl2.put("TOT_PO", "228");
        repl2.put("TOT_TA", "229");
        repl2.put("TOT_TB", "230");
        repl2.put("LASTQUARTERNAME", "231");
        repl2.put("CURRENTQUARTERNAME", "232");
        repl2.put("CLAIMSINCLUDING_LQBEP", "233");
        repl2.put("CLAIMSINCLUDING_CQBEP", "234");
        repl2.put("CLAIMSCERTIFIEDBYMORTGAGES_LQBEP", "133");
        repl2.put("CLAIMSCERTIFIEDBYMORTGAGES_CQBEP", "134");
        repl2.put("CLAIMSCONSTRUCTIONNOTCOMPLETED_LQBEP", "235");
        repl2.put("CLAIMSCONSTRUCTIONNOTCOMPLETED_CQBEP", "236");
        repl2.put("CLAIMSBYRESIDENTIALMORTGAGES_LQBEP", "237");
        repl2.put("CLAIMSBYRESIDENTIALMORTGAGES_CQBEP", "238");
        repl2.put("CLAIMSRESIDENTIALPREMISES_LQBEP", "239");
        repl2.put("CLAIMSRESIDENTIALPREMISES_CQBEP", "240");
        repl2.put("PARTICIPATIONCERTIFICATES_LQBEP", "241");
        repl2.put("PARTICIPATIONCERTIFICATES_CQBEP", "242");
        repl2.put("CASHINCLUDING_LQBEP", "243");
        repl2.put("CASHINCLUDING_CQBEP", "244");
        repl2.put("CURRENCYOFRF_LQBEP", "245");
        repl2.put("CURRENCYOFRF_CQBEP", "246");
        repl2.put("FOREIGNCURRENCY_LQBEP", "247");
        repl2.put("FOREIGNCURRENCY_CQBEP", "248");
        repl2.put("SIZEMORTGAGECOVERAGETOTAL_LQBEP", "249");
        repl2.put("SIZEMORTGAGECOVERAGETOTAL_CQBEP", "250");
        repl2.put("ESTIMATEDCCOSTPERISU_LQBEP", "251");
        repl2.put("ESTIMATEDCOSTPERISU_CQBEP", "252");

        Map<String,Object> repl3 = new HashMap();
        repl3.put("P1_PCI", "31");
        repl3.put("P1_AI", "32");
        repl3.put("P1_AIO", "33");
        repl3.put("P2_PCI", "34");
        repl3.put("P2_AI", "35");
        repl3.put("P2_AIO", "36");
        repl3.put("P3_PCI", "37");
        repl3.put("P3_AI", "38");
        repl3.put("P3_AIO", "39");
        repl3.put("P4_PCI", "310");
        repl3.put("P4_AI", "311");
        repl3.put("P4_AIO", "312");
        repl3.put("TOT_PCI", "313");
        repl3.put("TOT_AI", "314");
        repl3.put("TOT_AIO", "315");
        repl3.put("P1_PO", "316");
        repl3.put("P1_TA", "317");
        repl3.put("P1_TB", "318");
        repl3.put("P2_PO", "319");
        repl3.put("P2_TA", "320");
        repl3.put("P2_TB", "321");
        repl3.put("P3_PO", "322");
        repl3.put("P3_TA", "323");
        repl3.put("P3_TB", "324");
        repl3.put("P4_PO", "325");
        repl3.put("P4_TA", "326");
        repl3.put("P4_TB", "327");
        repl3.put("TOT_PO", "328");
        repl3.put("TOT_TA", "329");
        repl3.put("TOT_TB", "330");
        repl3.put("LASTQUARTERNAME", "331");
        repl3.put("CURRENTQUARTERNAME", "332");
        repl3.put("CLAIMSINCLUDING_LQBEP", "333");
        repl3.put("CLAIMSINCLUDING_CQBEP", "334");
        repl3.put("CLAIMSCERTIFIEDBYMORTGAGES_LQBEP", "133");
        repl3.put("CLAIMSCERTIFIEDBYMORTGAGES_CQBEP", "134");
        repl3.put("CLAIMSCONSTRUCTIONNOTCOMPLETED_LQBEP", "335");
        repl3.put("CLAIMSCONSTRUCTIONNOTCOMPLETED_CQBEP", "336");
        repl3.put("CLAIMSBYRESIDENTIALMORTGAGES_LQBEP", "337");
        repl3.put("CLAIMSBYRESIDENTIALMORTGAGES_CQBEP", "338");
        repl3.put("CLAIMSRESIDENTIALPREMISES_LQBEP", "339");
        repl3.put("CLAIMSRESIDENTIALPREMISES_CQBEP", "340");
        repl3.put("PARTICIPATIONCERTIFICATES_LQBEP", "341");
        repl3.put("PARTICIPATIONCERTIFICATES_CQBEP", "342");
        repl3.put("CASHINCLUDING_LQBEP", "343");
        repl3.put("CASHINCLUDING_CQBEP", "344");
        repl3.put("CURRENCYOFRF_LQBEP", "345");
        repl3.put("CURRENCYOFRF_CQBEP", "346");
        repl3.put("FOREIGNCURRENCY_LQBEP", "347");
        repl3.put("FOREIGNCURRENCY_CQBEP", "348");
        repl3.put("SIZEMORTGAGECOVERAGETOTAL_LQBEP", "349");
        repl3.put("SIZEMORTGAGECOVERAGETOTAL_CQBEP", "350");
        repl3.put("ESTIMATEDCCOSTPERISU_LQBEP", "351");
        repl3.put("ESTIMATEDCOSTPERISU_CQBEP", "352");


        sectionDataList.add(new SectionData(repl1));
        sectionDataList.add(new SectionData(repl2));
        sectionDataList.add(new SectionData(repl3));


        GenerateSectionsWordImpl generateSectionsWord =  new GenerateSectionsWordImpl();
        WordprocessingMLPackage template =  generateSectionsWord.getTemplate("C:\\Users\\Серяковы\\IdeaProjects\\docx\\src\\main\\resources\\templates\\Заключение АЗ (УК).docx");

        List<Object> listElement =  generateSectionsWord.getListDeepCopyElementFromTemplate(template,"<StartSection>","<EndSection>");

        listElement = generateSectionsWord.generateElementBySectionData(sectionDataList, listElement);

        generateSectionsWord.insertGeneratedElementIntoTemplate(template, listElement);

       generateSectionsWord.writeDocxToStream(template,"template2.docx");


    }


}
