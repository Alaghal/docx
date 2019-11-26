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
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class testWorkTabl {
	
	private int startIndexForInsertSection = -1;
	private int endIndexSection = -1;
	
	public WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
		WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
		return template;
	}
	
	public void writeDocxToStream(WordprocessingMLPackage template, String target) throws Docx4JException {
		File f = new File(target);
		template.save(f);
	}
	
	public void writeDocxToStream(WordprocessingMLPackage template, OutputStream target) throws Docx4JException {
		template.save(target);
	}
	
	public void fillTemplateBySanctionsData(WordprocessingMLPackage template, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder, String startSeparateElement, String endSeparateElement) throws JAXBException, XPathBinderAssociationIsPartialException {
		
		List<Object> listElement = getListDeepCopyElementFromTemplate(template, startGenerateSectionPlaceholder, endGenerateSectionPlaceholder);
		
		listElement = generateElementBySectionData(sectionDataListData, listElement, startSeparateElement, endSeparateElement);
		
		insertGeneratedElementIntoTemplate(template, listElement);
		
	}
	
	private List<Object> getListDeepCopyElementFromTemplate(WordprocessingMLPackage template, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, XPathBinderAssociationIsPartialException {
		List<Object> elementForGenerate = new ArrayList();
		List<Object> elementDocument = template.getMainDocumentPart().getContent();
		startIndexForInsertSection = getIndexTextElement(template, startGenerateSectionPlaceholder);
		endIndexSection = getIndexTextElement(template, endGenerateSectionPlaceholder);
		
		for (int i = startIndexForInsertSection; i < endIndexSection; i++) {
			elementForGenerate.add(XmlUtils.deepCopy(elementDocument.get(i)));
		}
		return elementForGenerate;
	}
	
	private List<Object> generateElementBySectionData(List<SectionData> sectionDataListData, List<Object> elementForGenerate, String startSeparateElement, String endSeparateElement) {
		List<Object> resultList = new ArrayList();
		for (SectionData sectionData : sectionDataListData) {
			resultList.addAll(fillGeneratedSanctions(elementForGenerate, sectionData.getDataToAddForParagraps(), sectionData.getDataToAddForTable(), startSeparateElement, endSeparateElement));
		}
		return resultList;
	}
	
	private void insertGeneratedElementIntoTemplate(WordprocessingMLPackage template, List<Object> resultList) {
		if (startIndexForInsertSection > 0 && endIndexSection > 0 && !resultList.isEmpty()) {
			int startSection = 1;
			int endSection = 1;
			int lengthForRemove = endIndexSection - startSection + endSection;
			int startForRemove = startIndexForInsertSection - 1;
			for (int i = startForRemove; i <= lengthForRemove; i++) {
				template.getMainDocumentPart().getContent().remove(startForRemove);
			}
			template.getMainDocumentPart().getContent().addAll(startIndexForInsertSection, resultList);
		}
	}
	
	private int getIndexTextElement(WordprocessingMLPackage template, String textForSearch) throws JAXBException, XPathBinderAssociationIsPartialException {

		final String XPATH_TO_SELECT_TEXT_NODES = "//w:p";
		final List<Object> jaxbNodes = template.getMainDocumentPart().getJAXBNodesViaXPath(XPATH_TO_SELECT_TEXT_NODES, false);
		List<String> list = jaxbNodes.stream().map(Object::toString).collect(Collectors.toList());

		return list.indexOf(textForSearch);



	}
	
	private int getIndexFromCustomParagraphsTextElements(WordprocessingMLPackage template, Text text, String textForSearch) {
		int index = -1;
		Object objParentText = text.getParent();
		if (objParentText instanceof R) {
			Object objParenrt = ((R) objParentText).getParent();
			if (objParenrt instanceof P) {
				List<Object> paragraphElements = ((P) objParenrt).getContent();
				StringBuilder stringBuilder = new StringBuilder();
				
				for (Object paragrapsObject : paragraphElements) {
					if (paragrapsObject instanceof R) {
						List<Object> objectList = ((R) paragrapsObject).getContent();
						for (Object elementR : objectList) {
							
							Text textElement = (Text) ((JAXBElement) elementR).getValue();
							stringBuilder.append(textElement.getValue());
							
						}
					}
				}
				
				if (stringBuilder.toString().equals(textForSearch)) {
					index = template.getMainDocumentPart().getContent().indexOf(objParenrt);
					return index;
				}
			}
		}
		return index;
	}
	
	private List<Object> fillGeneratedSanctions(List<Object> elementDocument, Map<String, String> dataForReplace, List<Map<String, String>> dataToAddForTable, String startSeparateElement, String endSeparateElement) {
		List<Object> cloneElementDocument = cloneList(elementDocument);
		for (int i = 0; i < cloneElementDocument.size(); i++) {
			Object obj = cloneElementDocument.get(i);
			
			if (dataForReplace != null && obj instanceof P) {
				replaceParagraphsText(obj, dataForReplace, startSeparateElement, endSeparateElement);
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
	
	private void replaceParagraphsText(Object obj, Map<String, String> dataForReplace, String startSeparateElement, String endSeparateElement) {
		List<Object> paragraphElements = ((P) obj).getContent();
		StringBuilder stringBuilder = new StringBuilder();
		List<Object> listForRemove = new ArrayList<>();
		boolean rangePlaceholderText = false;
		
		for (Object paragrapsObject : paragraphElements) {
			if (paragrapsObject instanceof R) {
				List<Object> objectList = ((R) paragrapsObject).getContent();
				for (Object elementR : objectList) {
					
					Text textElement = (Text) ((JAXBElement) elementR).getValue();
					
					if (textElement.getValue().contains(startSeparateElement) && textElement.getValue().contains(endSeparateElement)) {
						String replacementValue = dataForReplace.get(textElement.getValue());
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
						String replacementOfCustomValue = dataForReplace.get(stringBuilder.toString());
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
		
		((P) obj).getContent().removeAll(listForRemove);
	}
	
	private void replaceTableData(Object obj, List<Map<String, String>> dataToAddForTable) {
		Tbl table = (Tbl) ((JAXBElement) obj).getValue();
		List<Object> rows = getAllElementFromObject(table, Tr.class);
		for (int j = 0, k = 1; j < rows.size() - 1; j++, k++) {
			
			Tr templateRow = (Tr) rows.get(k);
			Map<String, String> replacements = dataToAddForTable.get(j);
			
			addRowToTable(table, templateRow, replacements);
			
			table.getContent().remove(templateRow);
		}
	}
	
	public static void addRowToTable(Tbl reviewtable, Tr templateRow, Map<String, String> replacements) {
		Tr workingRow = XmlUtils.deepCopy(templateRow);
		List textElements = getAllElementFromObject(workingRow, Text.class);
		for (Object object : textElements) {
			Text text = (Text) object;
			String replacementValue = replacements.get(text.getValue());
			if (replacementValue != null)
				text.setValue(replacementValue);
		}
		
		reviewtable.getContent().add(workingRow);
	}
	
	private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
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
	
	private P createPageBreak() {
		Br br = Context.getWmlObjectFactory().createBr();
		br.setType(STBrType.PAGE);
		
		R run = Context.getWmlObjectFactory().createR();
		run.getContent().add(br);
		
		P paragraph = Context.getWmlObjectFactory().createP();
		paragraph.getContent().add(run);
		return paragraph;
	}
}
