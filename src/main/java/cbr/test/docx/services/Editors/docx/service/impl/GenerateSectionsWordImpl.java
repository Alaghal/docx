package cbr.test.docx.services.Editors.docx.service.impl;

import cbr.test.docx.services.Editors.docx.component.SectionData;
import cbr.test.docx.services.Editors.docx.service.GenerateSectionsWord;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
//import ru.atc.uncrm.docx4j.component.SectionData;
//import ru.atc.uncrm.docx4j.service.GenerateSectionsWord;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


public class GenerateSectionsWordImpl implements GenerateSectionsWord {
    private  int startIndexSection = -1;
    private  int endIndexSection = -1;
    private final String startSeparateElement  = "<%";
    private final String endSeparateElement = "%>";

    @Override
    public void process(ByteArrayOutputStream outputStream, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, Docx4JException {
        WordprocessingMLPackage template = getTemplate(outputStream);
        List<Object> listElement =  getListDeepCopyElementFromTemplate(template,startGenerateSectionPlaceholder,endGenerateSectionPlaceholder);

        listElement = generateElementBySectionData(sectionDataListData, listElement);

        insertGeneratedElementIntoTemplate(template, listElement);

        //template.getMainDocumentPart().addParagraphOfText("Hello Word by RCR!");

       // writeDocxToStream(template, outputStream);ераенкевкевапвпа
        writeDocxToStream(template,"curator.docx");

    }

    private WordprocessingMLPackage getTemplate(ByteArrayOutputStream outputStream) throws  Docx4JException {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new ByteArrayInputStream(outputStream.toByteArray()));
        return template;
    }

    public WordprocessingMLPackage getTemplate(String path) throws  Docx4JException {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new File(path));
        return template;
    }

    public void writeDocxToStream(WordprocessingMLPackage template, String target) throws Docx4JException {
        File f = new File(target);
        template.save(f);
    }

    private void writeDocxToStream(WordprocessingMLPackage template, ByteArrayOutputStream target) throws Docx4JException {
        template.save(target);
    }

    public List<Object> getListDeepCopyElementFromTemplate(WordprocessingMLPackage template, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, XPathBinderAssociationIsPartialException {
        List<Object> elementForGenerate = new ArrayList();
        List<Object> elementDocument = template.getMainDocumentPart().getContent();
        startIndexSection = elementDocument.stream().map(Object::toString).collect(Collectors.toList()).indexOf(startGenerateSectionPlaceholder);
        endIndexSection = elementDocument.stream().map(Object::toString).collect(Collectors.toList()).indexOf(endGenerateSectionPlaceholder);
        int nextIndexAfterStartIndexSection = startIndexSection + 1;

        for (int i = nextIndexAfterStartIndexSection; i < endIndexSection ; i++) {
            elementForGenerate.add(XmlUtils.deepCopy(elementDocument.get(i)));
        }

        return elementForGenerate;
    }

    public List<Object> generateElementBySectionData(List<SectionData> sectionDataListData, List<Object> elementForGenerate) {
        List<Object> resultList = new ArrayList();



        for (SectionData sectionData : sectionDataListData) {
            resultList.addAll(fillGeneratedSanctions(elementForGenerate, sectionData.getData()));
        }

        return resultList;
    }



    public void insertGeneratedElementIntoTemplate(WordprocessingMLPackage template, List<Object> resultList) {
        if (!resultList.isEmpty()) {
            for (int i = startIndexSection; i<= endIndexSection; i++){
                template.getMainDocumentPart().getContent().remove(startIndexSection);
            }
            template.getMainDocumentPart().getContent().addAll(startIndexSection, resultList);
        }
    }


    private List<Object> fillGeneratedSanctions(List<Object> elementDocument, Map<String, Object>data) {
        List<Object> cloneElementDocument = cloneList(elementDocument);
        for (int i = 0; i < cloneElementDocument.size(); i++) {
            Object obj = cloneElementDocument.get(i);

            if (obj instanceof P) {
                replaceParagraphsText(obj, data);
            } else if (obj instanceof JAXBElement) {
                replaceTableData(obj, data);
            }
        }

        return cloneElementDocument;
    }

    private List<Object> cloneList(List<Object> elementDocument) {
        List<Object> cloneList = new ArrayList(elementDocument.size());
        for (Object item : elementDocument) cloneList.add(XmlUtils.deepCopy(item));
        return cloneList;
    }


    private void replaceTableData(Object obj, Map<String, Object> dataToAddForTable) {
        Tbl table = (Tbl) ((JAXBElement) obj).getValue();
        List<Object> rows = getAllElementFromObject(table, Tr.class);
        for (int j = 0; j < rows.size(); j++) {

            Tr templateRow = (Tr) rows.get(j);

            addRowToTable(table, templateRow, dataToAddForTable);

            table.getContent().remove(templateRow);
        }

    }

    public  void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, Object> replacements) {
        Tr workingRow = XmlUtils.deepCopy(templateRow);
        List pElements = getAllElementFromObject(workingRow, P.class);
        for (Object object : pElements) {
            replaceParagraphsText(object,replacements);

        }

        reviewtable.getContent().add(workingRow);

    }

    private void replaceParagraphsText(Object obj, Map<String, Object> dataForReplace) {
        List<Object> paragraphElements = ((P) obj).getContent();
        StringBuilder stringBuilder = new StringBuilder();
        List<Object> listForRemove = new ArrayList<>();
        boolean rangePlaceholderText = false;

        for (Object paragrapsObject : paragraphElements) {
            if (paragrapsObject instanceof R) {
                List<Object> objectList = ((R) paragrapsObject).getContent();
                for (Object elementR : objectList) {

                    if (!((JAXBElement) elementR).getDeclaredType().getName().equals("org.docx4j.wml.R$LastRenderedPageBreak")){

                    Text textElement = (Text) ((JAXBElement) elementR).getValue();

                    if (textElement.getValue().contains(startSeparateElement) && textElement.getValue().contains(endSeparateElement)) {
                        String replacementValue = String.valueOf (dataForReplace.get(deleteSeparateElementIntoText(textElement.getValue())));
                        if (replacementValue != null) {
                            textElement.setValue(replacementValue);
                        }
                    } else if (textElement.getValue().contains(startSeparateElement)) {
                        stringBuilder.append(textElement.getValue());
                        listForRemove.add(textElement.getParent());
                        rangePlaceholderText = true;

                    } else if (textElement.getValue().contains(endSeparateElement)) {
                        rangePlaceholderText = false;
                        stringBuilder.append(textElement.getValue());
                        String replacementOfCustomValue = String.valueOf(dataForReplace.get(deleteSeparateElementIntoText(stringBuilder.toString())));
                        if (replacementOfCustomValue != null) {
                            textElement.setValue(replacementOfCustomValue);
                        }

                    } else if (rangePlaceholderText) {
                        stringBuilder.append(textElement.getValue());
                        listForRemove.add(textElement.getParent());
                    }
                    }
                }

            }
        }

        ((P) obj).getContent().removeAll(listForRemove);
    }

    private String deleteSeparateElementIntoText(String text){

        return text.replace(startSeparateElement,"").replace(endSeparateElement,"").trim();
    }


    private List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));

            }

        }

        return result;

    }

}
