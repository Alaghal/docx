package ru.atc.uncrm.docx4j.component;

import org.docx4j.openpackaging.exceptions.Docx4JException;

import javax.xml.bind.JAXBException;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;

public interface WriterDataSectionsForTemplates {

    OutputStream writeGenerateSection(ByteArrayOutputStream outputStream) throws JAXBException, Docx4JException;
}
