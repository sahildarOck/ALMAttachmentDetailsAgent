package com.alm.wrapper.classes;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;

import atu.alm.wrapper.classes.TDConnection;

/**
 * Class to wrap the functionalities pertaining to ALM - OTA COM API
 * 
 * @author sahil.srivastava
 *
 */

public class ALMAutomationWrapper {

	// private static ALMServiceWrapper wrapper;
	private TDConnection tdConn;
	private ActiveXComponent almActiveXComponent;
	private ALMData almData;

	private String almURL;
	private String userName;
	private String password;
	private String domain;
	private String project;

	public ALMAutomationWrapper(ALMData almData) {
		this.almData = almData;

		almURL = ALMData.getAlmURL();
		userName = ALMData.getUserName();
		password = String.valueOf(ALMData.getPassword());
		domain = ALMData.getDomain();
		project = ALMData.getProject();
	}

	public boolean connectAndLoginALM() {
		almActiveXComponent = new ActiveXComponent("TDAPIOLE80.TDConnection");
		almActiveXComponent.invoke("InitConnectionEx", almURL);
		almActiveXComponent.invoke("Login", new Variant(userName), new Variant(password));
		almActiveXComponent.invoke("Connect", new Variant(domain), new Variant(project));
		return true;
	}

	public void closeConnection() {
		// wrapper.close();
		System.out.println("invoking log out");
		almActiveXComponent.invoke("Logout");
	}

	public ALMData getAlmData() {
		return almData;
	}

	public void setAlmData(ALMData almData) {
		this.almData = almData;
	}

	public TDConnection getTdConn() {
		return tdConn;
	}

	public void setTdConn(TDConnection tdConn) {
		this.tdConn = tdConn;
	}

	public ActiveXComponent getAlmActiveXComponent() {
		return almActiveXComponent;
	}

	public void setAlmActiveXComponent(ActiveXComponent almActiveXComponent) {
		this.almActiveXComponent = almActiveXComponent;
	}
}