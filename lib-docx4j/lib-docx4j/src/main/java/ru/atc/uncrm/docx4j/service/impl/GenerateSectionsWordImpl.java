package ru.atc.uncrm.docx4j.service.impl;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import ru.atc.uncrm.docx4j.component.SectionData;
import ru.atc.uncrm.docx4j.service.GenerateSectionsWord;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


public class GenerateSectionsWordImpl implements GenerateSectionsWord {
    private  int startIndexForInsertSection = -1;
    private  int endIndexSection = -1;
    private final String startSeparateElement  = "<%";
    private final String endSeparateElement = "%>";

    @Override
    public ByteArrayOutputStream process(ByteArrayOutputStream outputStream, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, Docx4JException {
        WordprocessingMLPackage template = getTemplate(outputStream);

        //List<Object> listElement =  getListDeepCopyElementFromTemplate(template,startGenerateSectionPlaceholder,endGenerateSectionPlaceholder);

        //listElement = generateElementBySectionData(sectionDataListData, listElement);

       // insertGeneratedElementIntoTemplate(template, listElement);

        template.getMainDocumentPart().addParagraphOfText("Hello Word by RCR!");

      return writeDocxToStream(template, outputStream);
    }

    private WordprocessingMLPackage getTemplate(ByteArrayOutputStream outputStream) throws  Docx4JException {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(new ByteArrayInputStream(outputStream.toByteArray()));
        return template;
    }



    private ByteArrayOutputStream writeDocxToStream(WordprocessingMLPackage template, ByteArrayOutputStream target) throws Docx4JException {

        template.save(target);
        return target;
    }

    private List<Object> getListDeepCopyElementFromTemplate(WordprocessingMLPackage template, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, XPathBinderAssociationIsPartialException {
        List<Object> elementForGenerate = new ArrayList();
        List<Object> elementDocument = template.getMainDocumentPart().getContent();
        startIndexForInsertSection = elementDocument.stream().map(Object::toString).collect(Collectors.toList()).indexOf(startGenerateSectionPlaceholder);
        endIndexSection = elementDocument.stream().map(Object::toString).collect(Collectors.toList()).indexOf(endGenerateSectionPlaceholder);

        for (int i = startIndexForInsertSection+1; i <= endIndexSection -1; i++) {
            elementForGenerate.add(XmlUtils.deepCopy(elementDocument.get(i)));
        }
        return elementForGenerate;
    }

    private List<Object> generateElementBySectionData(List<SectionData> sectionDataListData, List<Object> elementForGenerate) {
        List<Object> resultList = new ArrayList();
        for (SectionData sectionData : sectionDataListData) {
            resultList.addAll(fillGeneratedSanctions(elementForGenerate, sectionData.getDataToAddForParagraphs(), sectionData.getDataToAddForTable()));
        }
        return resultList;
    }

    private void insertGeneratedElementIntoTemplate(WordprocessingMLPackage template, List<Object> resultList) {
        if (startIndexForInsertSection > 0 && endIndexSection > 0 && !resultList.isEmpty()) {

            for (int i = startIndexForInsertSection; i <= endIndexSection; i++) {
                template.getMainDocumentPart().getContent().remove(i);
            }
            template.getMainDocumentPart().getContent().addAll(startIndexForInsertSection, resultList);
        }
    }


    private List<Object> fillGeneratedSanctions(List<Object> elementDocument, Map<String, Object> dataForParagraphsReplace, List<Map<String, Object>> dataToAddForTable) {
        List<Object> cloneElementDocument = cloneList(elementDocument);
        for (int i = 0; i < cloneElementDocument.size(); i++) {
            Object obj = cloneElementDocument.get(i);

            if (!dataForParagraphsReplace.isEmpty() && obj instanceof P) {
                replaceParagraphsText(obj, dataForParagraphsReplace);
            } else if (obj instanceof JAXBElement) {
                replaceTableData(obj, dataToAddForTable);
            }
        }

        return cloneElementDocument;
    }

    private List<Object> cloneList(List<Object> elementDocument) {
        List<Object> cloneList = new ArrayList(elementDocument.size());
        for (Object item : elementDocument) cloneList.add(XmlUtils.deepCopy(item));
        return cloneList;
    }


    private void replaceTableData(Object obj, List<Map<String, Object>> dataToAddForTable) {
        Tbl table = (Tbl) ((JAXBElement) obj).getValue();
        List<Object> rows = getAllElementFromObject(table, Tr.class);
        for (int j = 0; j < rows.size(); j++) {

            Tr templateRow = (Tr) rows.get(j);

            addRowToTable(table, templateRow, dataToAddForTable);

            table.getContent().remove(templateRow);
        }

    }

    public  void addRowToTable(Tbl reviewtable, Tr templateRow, List<Map<String, Object>> replacements) {
        Tr workingRow = XmlUtils.deepCopy(templateRow);
        List pElements = getAllElementFromObject(workingRow, P.class);
        for (Object object : pElements) {

            for (Map <String,Object> map:replacements) {

                replaceParagraphsText(object,map);
            }

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

        return text.replace(startSeparateElement,"").replace(endSeparateElement,"");
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
