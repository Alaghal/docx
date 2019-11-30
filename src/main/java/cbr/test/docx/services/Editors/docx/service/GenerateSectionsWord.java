package cbr.test.docx.services.Editors.docx.service;

import cbr.test.docx.services.Editors.docx.component.SectionData;
import org.docx4j.openpackaging.exceptions.Docx4JException;
//import ru.atc.uncrm.docx4j.component.SectionData;

import javax.xml.bind.JAXBException;
import java.io.ByteArrayOutputStream;
import java.util.List;

public interface GenerateSectionsWord {

    void process(ByteArrayOutputStream outputStream, List<SectionData> sectionDataListData, String startGenerateSectionPlaceholder, String endGenerateSectionPlaceholder) throws JAXBException, Docx4JException;
}
