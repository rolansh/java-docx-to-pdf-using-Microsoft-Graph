/*
 *  Copyright 2020, Plutext Pty Ltd.
 *   
    This module is licensed under the Apache License, Version 2.0 (the "License"); 
    you may not use this file except in compliance with the License. 

    You may obtain a copy of the License at 

        http://www.apache.org/licenses/LICENSE-2.0 

    Unless required by applicable law or agreed to in writing, software 
    distributed under the License is distributed on an "AS IS" BASIS, 
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
    See the License for the specific language governing permissions and 
    limitations under the License.

 */

package org.plutext.msgraph.convert.graphsdk;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.UUID;
import java.util.concurrent.CompletableFuture;
import java.util.stream.Stream;

import com.azure.core.http.rest.RequestOptions;
import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.authentication.TokenCredentialAuthProvider;
import com.microsoft.graph.logger.LoggerLevel;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.DriveItemCreateUploadSessionParameterSet;
import com.microsoft.graph.models.DriveItemUploadableProperties;
import com.microsoft.graph.options.Option;
import com.microsoft.graph.requests.DriveCollectionPage;
import com.microsoft.graph.requests.DriveCollectionRequestBuilder;
import com.microsoft.graph.requests.DriveItemContentStreamRequestBuilder;
import com.microsoft.graph.requests.DriveItemCreateUploadSessionRequest;
import com.microsoft.graph.requests.DriveItemRequest;
import com.microsoft.graph.requests.GraphServiceClient;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.plutext.msgraph.convert.AbstractOpenXmlToPDF;
import org.plutext.msgraph.convert.AuthConfig;
import org.plutext.msgraph.convert.ConversionException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Demonstrate using the Graph SDK high level API for PDF Conversion.
 * 
 * This doesn't support converting files bigger than 4MB, so you should use PdfConverterLarge instead.  
 * 
 *
 * @author jharrop
 *
 */
public class Limited4MB extends AbstractOpenXmlToPDF {

	public Limited4MB(AuthConfig authConfig) {
		super(authConfig);
	}

	private static final Logger log = LoggerFactory.getLogger(Limited4MB.class);
			
	@Override
	public byte[] convert(byte[] bytes, String ext) throws ConversionException {
		// Authentication provider - Client credentials provider - Using a client secret
		// https://learn.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=java
		final ClientSecretCredential clientCreds = new ClientSecretCredentialBuilder()
				.clientId(authConfig.apiKey())
				.clientSecret(authConfig.apiSecret())
				.tenantId(authConfig.tenant())
				.build();

		final TokenCredentialAuthProvider authProvider = new TokenCredentialAuthProvider(clientCreds);

		final GraphServiceClient graphClient = GraphServiceClient
				.builder()
				.authenticationProvider(authProvider)
				.buildClient();
		
        String tmpFileName = UUID.randomUUID().toString() + ".docx"; // an extension is required
//		String itemPath =  "root:/" + tmpFileName +":";

		final DriveItemCreateUploadSessionParameterSet uploadParams = DriveItemCreateUploadSessionParameterSet
				.newBuilder()
				.withItem(new DriveItemUploadableProperties())
				.build();

		// Upload
		DriveItem driveItem = graphClient
				.sites()
				.byId(authConfig.site())
				.drive()
				.items()
				.byId("root")
				.itemWithPath(tmpFileName)
				.content()
				.buildRequest()
				.put(bytes);

		log.debug("............... {}, {}, {}, {}", driveItem.id, driveItem.webUrl, driveItem.webDavUrl, driveItem.folder);
		// Download as pdf
//		String customPdfUrlStr = graphClient.sites(authConfig.site()).drive().items("root:/" + tmpFileName + ":")
//				.buildRequest().getRequestUrl().toString()
//				;
//				+ "/content?format=pdf";


		graphClient.getLogger().setLoggingLevel(LoggerLevel.DEBUG);
		String customPdfUrlStr = "/sites/" + authConfig.site() + "/drive/items/root:/" + tmpFileName + ":" + "/content?format=pdf";

		BufferedInputStream inputStream = (BufferedInputStream)graphClient.customRequest(customPdfUrlStr, Stream.class)
				.buildRequest()
				.get();



		// Method 2
		DriveItemContentStreamRequestBuilder pdfRequestBuilder = graphClient
				.sites()
				.byId(authConfig.site())
				.drive()
				.items()
				.byId("root")
				.itemWithPath(tmpFileName)
				.content();

		String pdfUrlStr = pdfRequestBuilder
				.buildRequest()
				.getRequestUrl()
				.toString() + "?format=pdf";

		// TODO: How to request the full path url?

		try {
			// Move to trash
			graphClient.sites(authConfig.site()).drive().items("root:/" + tmpFileName + ":")
					.buildRequest().delete();

			return IOUtils.toByteArray(inputStream);
		} catch (Exception e) {
			throw new ConversionException(e.getMessage(), e);
		}
		
	}

	public byte[] convert(File docx) throws ConversionException, IOException {
		
		String filename = docx.getName();
		String ext = filename.substring(filename.lastIndexOf("."));
		
		return convert( FileUtils.readFileToByteArray(docx), ext);
	}

	@Override
	public byte[] convert(InputStream docx, String ext) throws ConversionException, IOException {
		
		return convert( IOUtils.toByteArray(docx), ext );
	}	
	
	

}
