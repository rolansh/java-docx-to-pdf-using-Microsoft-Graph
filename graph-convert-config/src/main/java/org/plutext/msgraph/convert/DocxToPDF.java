package org.plutext.msgraph.convert;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

public abstract class DocxToPDF implements OpenXmlToPDF {
	
	public abstract byte[] convert(byte[] docx) throws IOException;  
	
	public abstract byte[] convert(File docx) throws IOException;  

	public abstract byte[] convert(InputStream docx) throws IOException;  

}
