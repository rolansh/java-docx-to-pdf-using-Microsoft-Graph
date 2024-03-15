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
import java.lang.reflect.InvocationTargetException;
import java.net.URL;
import java.util.UUID;
import java.util.concurrent.CancellationException;
import java.util.stream.Stream;

import com.google.gson.JsonElement;
import com.google.gson.JsonPrimitive;
import com.microsoft.graph.http.GraphServiceException;
import com.microsoft.graph.models.DriveItemCreateUploadSessionParameterSet;
import com.microsoft.graph.models.Request;
import com.microsoft.graph.requests.DriveItemContentStreamRequestBuilder;
import com.microsoft.graph.requests.DriveItemCreateUploadSessionRequest;
import com.microsoft.graph.requests.GraphServiceClient;
import com.microsoft.graph.tasks.IProgressCallback;
import com.microsoft.graph.tasks.LargeFileUploadResult;
import com.microsoft.graph.tasks.LargeFileUploadTask;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.plutext.msgraph.convert.AbstractOpenXmlToPDF;
import org.plutext.msgraph.convert.AuthConfig;
import org.plutext.msgraph.convert.ConversionException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.DriveItemUploadableProperties;
import com.microsoft.graph.models.UploadSession;

import com.microsoft.graph.authentication.TokenCredentialAuthProvider;







/**
 * Demonstrate using the Graph SDK high level API for PDF Conversion.
 * 
 * Supports converting large files, but won't update your TOC before doing so!
 * 
 * Vote for this enhancement at 
 * https://microsoftgraph.uservoice.com/forums/920506-microsoft-graph-feature-requests/suggestions/41235295-docx-to-pdf-file-conversion-update-table-of-conte 
 * 
 * @author jharrop
 *
 */
public class PdfConverterLarge  extends AbstractOpenXmlToPDF {

	public PdfConverterLarge(AuthConfig authConfig) {
		super(authConfig);
	}


	private static final Logger log = LoggerFactory.getLogger(PdfConverterLarge.class);


	public byte[] convert(InputStream fileStream, long streamSize, String ext) throws ConversionException, IOException {
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

		String tmpFileName = UUID.randomUUID().toString() + ext;

		try {
			DriveItemUploadableProperties properties = new DriveItemUploadableProperties();
			JsonPrimitive conflictBehavior = new JsonPrimitive("replace");
			properties.additionalDataManager().put("@microsoft.graph.conflictBehavior", conflictBehavior);

			final DriveItemCreateUploadSessionParameterSet uploadParams = DriveItemCreateUploadSessionParameterSet
					.newBuilder()
					.withItem(new DriveItemUploadableProperties())
					.build();
			final DriveItemCreateUploadSessionRequest uploadSessionRequest = graphClient
					.sites()
					.byId(authConfig.site())
					.drive()
					.items()
					.byId("root")
					.itemWithPath(tmpFileName)
					.createUploadSession(uploadParams)
					.buildRequest();

			log.info("uploadFile: uploading file to SharePoint DMS with request url : {}", uploadSessionRequest.getRequestUrl());

			final UploadSession uploadSession = uploadSessionRequest.post();
			if (uploadSession == null) {
				throw new ConversionException("uploadFile: failed to upload file to SharePoint DMS : nullable upload file session");
			}

			int maxSliceSize = 10 * 320 * 1024;

			final LargeFileUploadTask<DriveItem> fileUploadTask = new LargeFileUploadTask<>(
					uploadSession,
					graphClient,
					fileStream,
					streamSize,
					DriveItem.class
			);
			final IProgressCallback progressCallback = (current, max) -> log.info("uploadFile: uploaded {} bytes of {} total bytes", current, max);

			final LargeFileUploadResult<DriveItem> uploadFileResult = fileUploadTask.upload(0, null, progressCallback);

			log.info("uploadFile: uploaded '{}' file to SharePoint DMS", uploadFileResult.location);

			log.info("upload: file was successfully uploaded to SharePoint");

			// Download as pdf
			String customPdfUrlStr = "/sites/" + authConfig.site() + "/drive/items/root:" + tmpFileName + ":" + "/content?format=pdf";
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

			log.debug("compare:");
			log.debug("compare:");
			log.debug(pdfUrlStr);
			URL pdfUrl = new URL(pdfUrlStr);

			try {
				// Move to trash
				graphClient.sites(authConfig.site()).drive().items("root:" + tmpFileName + ":")
						.buildRequest().delete();

				return IOUtils.toByteArray(inputStream);
			} catch (Exception e) {
				throw new ConversionException(e.getMessage(), e);
			}

		} catch (Exception e) {
			throw new ConversionException(e.getMessage(), e);
		}
	}

	@Override
	public byte[] convert(byte[] docx, String ext) throws ConversionException {
		
		InputStream fileStream = new ByteArrayInputStream(docx);

		try {
			return convert( fileStream,  docx.length, ext);
		} catch (IOException e) {
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
		// inefficient, but we need length
		return convert( IOUtils.toByteArray(docx), ext );
	}
}
