package cbr.test.docx.services.Editors.docx;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.util.List;

public interface WorkTemplateService {
	
	WordprocessingMLPackage getTemplate(String name);
	
	void writeDocxToStream(WordprocessingMLPackage template, String target);
	
	void fillTemplateBySanctionsData(WordprocessingMLPackage template, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder, String startSeparateElement, String endSeparateElement);
	
}
