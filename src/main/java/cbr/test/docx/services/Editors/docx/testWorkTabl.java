package cbr.test.docx.services.Editors.docx;

import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.DocumentModel;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.Parts;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;


import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class testWorkTabl {

    public WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
       WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
       // WordprocessingMLPackage template = WordprocessingMLPackage.load();
        return template;
    }

    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
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


    public void writeDocxToStream(WordprocessingMLPackage template, String target// OutputStream outputStream
    ) throws IOException, Docx4JException {
       // template.save(outputStream);
        File f = new File(target);
        template.save(f);
    }


    public void fillTemplateBySanctionsData(WordprocessingMLPackage template, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) {

        List<Object> elementDocument = template.getMainDocumentPart().getContent();
        boolean forCopyToGenerationSection = false;
        List<Object> elementForGenerate = new ArrayList();
        int indexForStartInsert = 0;

        for (int i = 0; i < elementDocument.size(); i++) {
            Object obj = XmlUtils.unwrap(elementDocument.get(i));
            if (forCopyToGenerationSection) {
                if (obj instanceof Text) {
                    if (((Text) obj).getValue().equals(endGenerateSectionPlaceholder)) {
                        break;
                    }
                }
                elementForGenerate.add(XmlUtils.deepCopy(elementDocument.get(i)));
            }
            if (obj instanceof Text) {
                if (((Text) obj).getValue().equals(startGenerateSectionPlaceholder)) {
                    forCopyToGenerationSection = true;
                    indexForStartInsert = i;
                }
            }
        }
        List<Object> resultList = new ArrayList();

        for (SectionData sectionData : sectionDataListData) {
            resultList.addAll(fillGeneratedSactions(elementForGenerate, sectionData.getDataToAddForParagraps(), sectionData.getDataToAddForTable()));
        }

        if (indexForStartInsert > 0 && !resultList.isEmpty()) {
            template.getMainDocumentPart().getContent().addAll(indexForStartInsert, resultList);
        }

    }

    private List<Object> fillGeneratedSactions(List<Object> elementDocument, Map<String, String> dataForReplace, List<Map<String, String>> dataToAddForTable) {
        List<Object> cloneElementDocument = new ArrayList(elementDocument.size());
        for (Object item : cloneElementDocument) cloneElementDocument.add(XmlUtils.deepCopy(item));

        for (int i = 0; i < cloneElementDocument.size(); i++) {
            Object obj = XmlUtils.unwrap(cloneElementDocument.get(i));
            if ( !dataForReplace.isEmpty() && obj instanceof Text) {
                Text textElement = (Text) obj;
                String replacementValue = dataForReplace.get(textElement.getValue());
                if (replacementValue != null) {
                    textElement.setValue(replacementValue);
                }
            } else if (obj instanceof Tbl) {
                Tbl table = (Tbl) obj;
                List<Object> rows = getAllElementFromObject(table, Tr.class);
                for (int j = 1; j < rows.size(); j++) {
                    Tr templateRow = (Tr) rows.get(j);

                    for (Map<String, String> replacements : dataToAddForTable) {
                        addRowToTable(table, templateRow, replacements);
                    }

                    table.getContent().remove(templateRow);
                }
            }
        }
        return cloneElementDocument;
    }


    public static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements) {
        Tr workingRow = (Tr) XmlUtils.deepCopy(templateRow);
        List textElements = getAllElementFromObject(workingRow, Text.class);
        for (Object object : textElements) {
            Text text = (Text) object;
            String replacementValue = (String) replacements.get(text.getValue());
            if (replacementValue != null)
                text.setValue(replacementValue);
        }

        reviewtable.getContent().add(workingRow);
    }

    public P createPageBreak() {
        Br br = Context.getWmlObjectFactory().createBr();
        br.setType(STBrType.PAGE);

        R run = Context.getWmlObjectFactory().createR();
        run.getContent().add(br);

        P paragraph = Context.getWmlObjectFactory().createP();
        paragraph.getContent().add(run);
        return paragraph;
    }
}
