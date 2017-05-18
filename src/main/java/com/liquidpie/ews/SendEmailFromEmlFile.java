package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.DefaultExtendedPropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.MapiPropertyType;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.property.complex.EmailAddress;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import microsoft.exchange.webservices.data.property.complex.MimeContent;
import microsoft.exchange.webservices.data.property.definition.ExtendedPropertyDefinition;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;


public class SendEmailFromEmlFile {
	
	public static void main(String[] args) throws Exception {
		ExchangeService service = ConnectionUtil.getService();

		File dir = new File("folder/containing/eml(s)");
		File [] files = dir.listFiles(new FilenameFilter() {
		    @Override
		    public boolean accept(File dir, String name) {
		        return name.endsWith("email_file.eml");
		    }
		});


		for (File eml : files) {
			FileInputStream fileStream = null;
			try {
				fileStream = new FileInputStream(eml);

				byte[] buffer = new byte[(int) eml.length()];
				fileStream.read(buffer);

				EmailMessage msg = new EmailMessage(service);
				MimeContent mime = new MimeContent("UTF-8", buffer);
				msg.setMimeContent(mime);

				// Add custom header
				ExtendedPropertyDefinition headerElement = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "x-mycustom-header", MapiPropertyType.String);
				msg.setExtendedProperty(headerElement, "My custom header value");

				msg.save(new FolderId(WellKnownFolderName.Drafts));

				MessageBody messageBody = new MessageBody("");
				EmailAddress emailAddress = new EmailAddress("user@example.com");
				msg.forward(messageBody, emailAddress);
			}
			finally {
				fileStream.close();
			}
		}
	}
}
