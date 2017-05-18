package com.liquidpie.ews;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;

public class CreateFolder {
	
	public static void main(String[] args) throws Exception {
		
		ExchangeService service = ConnectionUtil.getService();
		
		Folder root = Folder.bind(service, WellKnownFolderName.MsgFolderRoot);
		Folder folder = new Folder(service);
		folder.setDisplayName("Test Folder");
		// creates the folder as a child of the ROOT folder.
		folder.save(root.getId());
	}
}