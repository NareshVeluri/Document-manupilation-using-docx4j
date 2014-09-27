package net.vsspl.samples;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Text;

public class DocsGen {
	
	
	private WordprocessingMLPackage getTemplate()
			throws Docx4JException, FileNotFoundException {
		WordprocessingMLPackage template = WordprocessingMLPackage
				.load(new FileInputStream(new File("/home/naresh/Desktop/template.docx")));
		return template;
	}

	private static List<Object> getAllElementFromObject(Object obj,
			Class<?> toSearch) {
		
		List<Object> result = new ArrayList<Object>();
		if (obj instanceof JAXBElement<?>)
			obj = ((JAXBElement<?>) obj).getValue();
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

	private void replacePlaceholder(WordprocessingMLPackage template,
			String name, String placeholder) {
		
		List<Object> texts = getAllElementFromObject(
				template.getMainDocumentPart(), Text.class);
		
		System.out.println(template.getMainDocumentPart().getContentType());
		for (Object text : texts) {
			Text textElement = (Text) text;
			System.out.println(textElement.getValue());
			if (textElement.getValue().equals(placeholder)) {
				textElement.setValue(name);
			}
		}
	}

	private void writeDocxToStream(WordprocessingMLPackage template
			) throws IOException, Docx4JException {
		File f = new File("/home/naresh/Desktop/template.docx");
		template.save(f);
	}
	
	public static void main(String[] args) throws Docx4JException, IOException {
		DocsGen obj = new DocsGen();
		
		WordprocessingMLPackage template = obj.getTemplate();
		
		obj.replacePlaceholder(template, "Naresh", "replace_name");
		obj.replacePlaceholder(template, "Male", "replace_gender");
		
		obj.writeDocxToStream(template);
		
		System.out.println("done writing");
	}
	
}
