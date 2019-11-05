package cbr.test.docx.services.Editors.docx;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
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
	
	
	public void fillTemplateBySanctionsData(WordprocessingMLPackage template, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, XPathBinderAssociationIsPartialException {
		
		List<Object> elementDocument = template.getMainDocumentPart().getContent();
		List<Object> elementForGenerate = new ArrayList();
		int startIndexSection = getIndexTextElement(template, startGenerateSectionPlaceholder) + 1;
		int endIndexSection = getIndexTextElement(template, endGenerateSectionPlaceholder);
		
		for (int i = startIndexSection; i < endIndexSection; i++) {
			elementForGenerate.add(XmlUtils.deepCopy(elementDocument.get(i)));
		}
		
		List<Object> resultList = new ArrayList();
		for (SectionData sectionData : sectionDataListData) {
			resultList.addAll(fillGeneratedSactions(elementForGenerate, sectionData.getDataToAddForParagraps(), sectionData.getDataToAddForTable()));
		}
		
		if (startIndexSection > 0 && !resultList.isEmpty()) {
			template.getMainDocumentPart().getContent().addAll(resultList);
		}
		
	}
	
	private int getIndexTextElement(WordprocessingMLPackage template, String textForSearch) throws JAXBException, XPathBinderAssociationIsPartialException {
		int index = -1;
		String textNodesXPath = "//w:t";
		List<Object> textNodes = template.getMainDocumentPart()
				.getJAXBNodesViaXPath(textNodesXPath, false);
		for (Object obj : textNodes) {
			Text text = (Text) ((JAXBElement) obj).getValue();
			String textValue = text.getValue();
			
			if (textValue.equals(textForSearch)) {
				Object objparent = text.getParent();
				if (objparent instanceof R) {
					Object objectParent2 = ((R) objparent).getParent();
					if (objectParent2 instanceof P) {
						index = template.getMainDocumentPart().getContent().indexOf(objectParent2);
						return index;
					}
					
				}
			}
		}
		
		return index;
	}
	
	private List<Object> cloneList(List<Object> elementDocument) {
		List<Object> cloneList = new ArrayList(elementDocument.size());
		for (Object item : elementDocument) cloneList.add(XmlUtils.deepCopy(item));
		return cloneList;
	}
	
	private List<Object> fillGeneratedSactions(List<Object> elementDocument, Map<String, String> dataForReplace, List<Map<String, String>> dataToAddForTable) {
		List<Object> cloneElementDocument = cloneList(elementDocument);
		for (int i = 0; i < cloneElementDocument.size(); i++) {
			Object obj = cloneElementDocument.get(i);
			
			if (dataForReplace !=null && obj instanceof P) {
				List<Object> paragraphs = ((P) obj).getContent();
				for (Object paragrapsObject : paragraphs) {
					if (paragrapsObject instanceof R) {
						List<Object> objectList = ((R) paragrapsObject).getContent();
						for (Object elementR : objectList) {
							if (elementR instanceof Text) {
								Text textElement = (Text) elementR;
								String replacementValue = dataForReplace.get(textElement.getValue());
								if (replacementValue != null) {
									textElement.setValue(replacementValue);
								}
							}
						}
					}
				}
				
			} else if (obj instanceof JAXBElement) {
				Tbl table = (Tbl) ((JAXBElement) obj).getValue();
				List<Object> rows = getAllElementFromObject(table, Tr.class);
				for (int j = 0, k = 1; j < rows.size()-1; j++, k++) {
					
					Tr templateRow = (Tr) rows.get(k);
					Map<String, String> replacements = dataToAddForTable.get(j);
					
					addRowToTable(table, templateRow, replacements);
					
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
