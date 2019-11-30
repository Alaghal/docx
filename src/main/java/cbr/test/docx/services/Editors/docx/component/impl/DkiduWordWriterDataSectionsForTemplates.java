//package cbr.test.docx.services.Editors.docx.component.impl;
//
//
//import cbr.test.docx.services.Editors.docx.component.SectionData;
//import org.docx4j.openpackaging.exceptions.Docx4JException;
//import ru.atc.template.model.TemplateModel;
//import ru.atc.uncrm.docx4j.component.SectionData;
//import ru.atc.uncrm.docx4j.component.WriterDataSectionsForTemplates;
//import ru.atc.uncrm.docx4j.service.DataParser;
//import ru.atc.uncrm.docx4j.service.GenerateSectionsWord;
//import ru.atc.uncrm.docx4j.service.impl.DataParserImpl;
//import ru.atc.uncrm.docx4j.service.impl.GenerateSectionsWordImpl;
//
//import javax.xml.bind.JAXBException;
//import java.io.ByteArrayOutputStream;
//import java.io.OutputStream;
//import java.util.*;
//
//
//public class DkiduWordWriterDataSectionsForTemplates implements WriterDataSectionsForTemplates {
//
//    private List<String> expressionTables;
//    private List<String> expressionParagraphs;
//    private TemplateModel templateModel;
//    private String startGenerateSectionPlaceholder;
//    private String endGenerateSectionPlaceholder;
//    private String separateKeyElement;
//
//    public DkiduWordWriterDataSectionsForTemplates(String codeReport, TemplateModel model) {
//
//        expressionParagraphs = new ArrayList<>();
//        this.templateModel = model;
//
//        //Добавление expression
//        switch (codeReport) {
//            case "AZ_UK_DKIDU":
//                expressionTables = Arrays.asList("/infoinv-isu/isunv_list","/infopay-isu/isupay_list","infoxyz_isu/isu_two_quarter");
//                startGenerateSectionPlaceholder = "<StartSection>";
//                endGenerateSectionPlaceholder = "<EndSection>";
//                separateKeyElement = "NAME";
//
//        }
//
//    }
//
//    private List<SectionData> getDataFromModel(){
//        List<SectionData> sectionDataList = new ArrayList<>();
//        Map<String,Object> dataForParagraphs = new HashMap<>();
//        DataParser dataParser =  new DataParserImpl();
//
//
//        if(!expressionParagraphs.isEmpty()) {
//            for (String expression : expressionParagraphs) {
//                Object value = dataParser.getMapFromTemplateModel(expression, templateModel);
//                dataForParagraphs.put(expression, value);
//            }
//        }
//
//        Map <String,  Map<String,Object>> mapingData  =  new HashMap<>();
//        Set<String> uniqDataKeyList = new HashSet<>();
//
//        for (String expressionTable:expressionTables ) {
//            List<Map<String,Object>> dataFromExpression = dataParser.getListMapFromTemplateModel(expressionTable, templateModel);
//
//            for (Map<String,Object> map :dataFromExpression) {
//                String value = (String) map.get(separateKeyElement);
//
//                if(value != null){
//
//                    if(uniqDataKeyList.contains(value)){
//                        mapingData.get(value).putAll(map);
//                    } else {
//                        uniqDataKeyList.add(value);
//                        mapingData.put(value,map);
//                    }
//                }
//            }
//
//        }
//
//        for (String key: mapingData.keySet()) {
//            sectionDataList.add(new SectionData(mapingData.get(key), dataForParagraphs));
//        }
//
//        return sectionDataList;
//    }
//
//
//    @Override
//    public OutputStream writeGenerateSection(ByteArrayOutputStream outputStream) throws JAXBException, Docx4JException {
//
//        List<SectionData> dataList = getDataFromModel();
//
//      //  GenerateSectionsWord generateSectionsWord =  new GenerateSectionsWordImpl();
//
//        generateSectionsWord.process(outputStream,dataList,startGenerateSectionPlaceholder,endGenerateSectionPlaceholder);
//        return outputStream;
//    }
//}
