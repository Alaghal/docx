package ru.atc.uncrm.docx4j.factory;

import ru.atc.template.model.TemplateModel;
import ru.atc.uncrm.docx4j.component.WriterDataSectionsForTemplates;

import java.io.OutputStream;
import java.util.Set;


public interface  FactoryTemplateWriterDataSection {

    WriterDataSectionsForTemplates createTemplateWriteDataSection();

}
