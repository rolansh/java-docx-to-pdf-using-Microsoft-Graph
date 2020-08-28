package org.plutext.msgraph.convert.msal;


import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.util.UUID;
import java.util.concurrent.ExecutionException;

import org.apache.commons.io.IOUtils;
import org.plutext.msgraph.convert.DocxToPdfConverter;
import org.plutext.msgraph.convert.AuthConfig;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.scribejava.core.httpclient.HttpClient;
import com.microsoft.aad.msal4j.ClientCredentialFactory;
import com.microsoft.aad.msal4j.ConfidentialClientApplication;


public class PdfConverter  extends DocxToPdfConverter  {
	
	/**
	 * PdfConverter using scribe's JDKHttpClient
	 * @param authConfig
	 */
	public PdfConverter(AuthConfig authConfig) {
		super(authConfig);
		fs = new FileService(getConfidentialClientApplication()); 
	}

	/**
	 * PdfConverter using specified HttpClient, configured in your pom.
	 * @param authConfig
	 */
	public PdfConverter(AuthConfig authConfig, HttpClient httpClient) {
		super(authConfig);
		fs = new FileService(getConfidentialClientApplication(), httpClient); 
		
	}
	
	private ConfidentialClientApplication getConfidentialClientApplication() {

		// See https://docs.microsoft.com/en-us/azure/active-directory/azuread-dev/v1-oauth2-client-creds-grant-flow
    	ConfidentialClientApplication confidentialClientApp =null;
		try {
			confidentialClientApp = ConfidentialClientApplication
			          .builder(authConfig.apiKey(), 
			        		  ClientCredentialFactory.createFromSecret(authConfig.apiSecret()))
			          //.authority("https://login.microsoftonline.com/common/oauth2/token") 
			          .authority("https://login.microsoftonline.com/" + authConfig.tenant() + "/oauth2/token")
			          .build();
		} catch (MalformedURLException e) {
			// shouldn't happen
			e.printStackTrace();
		}    	

		return confidentialClientApp;
	}

	private static final Logger LOG = LoggerFactory.getLogger(PdfConverter.class);
		
	FileService fs = null ;
	
	@Override
	public byte[] convert(byte[] docx) {
		try {
			
			// Upload the file
			// Let's work with a known filename.  This way we can ignore the returned itemid (which we need JSON parsing to read)
	        String tmpFileName = UUID.randomUUID()+ ".docx"; // TODO dotx/dotm etc
			String item =  "root:/" + tmpFileName +":";	
			String path = "https://graph.microsoft.com/v1.0/sites/" + authConfig.site() + "/drive/items/" + item + "/content";
			
			
			// Upload the file
			Boolean result = fs.uploadStreamAsync(path, docx, 
					"application/vnd.openxmlformats-officedocument.wordprocessingml.document").get();
			if (result==null || result.booleanValue()==false) {
				System.out.println("Upload failed, terminating.");
				throw new RuntimeException("upload failed");
			}
			
			// Convert
			byte[] pdfBytes = fs.downloadConvertedFileAsync(path + "?format=pdf").get();
			
			// Move temp file to recycle bin
			path = "https://graph.microsoft.com/v1.0/sites/" + authConfig.site() + "/drive/items/" + item;  // filename is easier than item id here			
			boolean deleted = fs.deleteFileAsync(path).get();
			System.out.println(deleted);
			
			return pdfBytes;
			
		} catch (InterruptedException e) {
			e.printStackTrace();
		} catch (ExecutionException e) {
			e.printStackTrace();
		}
		return null;
	}


	@Override
	public byte[] convert(InputStream docx) throws IOException {
		return convert( IOUtils.toByteArray(docx) );
	}	
	
	
	@Override
	public byte[] convert(File inFile) throws IOException {

		try {
			
			// Upload the file
			// Let's work with a known filename.  This way we can ignore the returned itemid (which we need JSON parsing to read)
	        String tmpFileName = UUID.randomUUID()+ ".docx"; // TODO dotx/dotm etc
			String item =  "root:/" + tmpFileName +":";	
			String path = "https://graph.microsoft.com/v1.0/sites/" + authConfig.site() + "/drive/items/" + item + "/content";
			
			
			// Upload the file
			Boolean result = fs.uploadStreamAsync(path, inFile, 
					"application/vnd.openxmlformats-officedocument.wordprocessingml.document").get();
			if (result==null || result.booleanValue()==false) {
				System.out.println("Upload failed, terminating.");
				throw new RuntimeException("upload failed");
			}
			
			// Convert
			byte[] pdfBytes = fs.downloadConvertedFileAsync(path + "?format=pdf").get();
			
			// Move temp file to recycle bin
			path = "https://graph.microsoft.com/v1.0/sites/" + authConfig.site() + "/drive/items/" + item;  // filename is easier than item id here			
			boolean deleted = fs.deleteFileAsync(path).get();
			System.out.println(deleted);
			
			return pdfBytes;
			
		} catch (InterruptedException e) {
			e.printStackTrace();
		} catch (ExecutionException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	
}