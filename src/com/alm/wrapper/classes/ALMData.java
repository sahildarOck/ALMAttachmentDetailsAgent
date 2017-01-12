package com.alm.wrapper.classes;

import java.io.File;

import atu.alm.wrapper.enums.StatusAs;

/**
 * Central class to store static data and data received from UI
 * 
 * @author sahil.srivastava
 *
 */
public class ALMData {

	private static String almURL;
	private static String userName;
	private static char[] password;
	private static String domain;
	private static String project;

	private int testSetID;
	private String testCaseName;

	private StatusAs testCaseStatus;

	private String defectID;
	private String failedTestStepNo;
	private String failedTestStepActualResult;

	private File testStepsXLFile;
	private File attachmentFile;
	
	private String testFolderPathOrTestSetID;
	private File outputXLFileLoc;

	private String runName;

	public static String getAlmURL() {
		return almURL;
	}

	public static void setAlmURL(String almURL) {
		ALMData.almURL = almURL;
	}

	public static String getUserName() {
		return userName;
	}

	public static void setUserName(String userName) {
		ALMData.userName = userName;
	}

	public static char[] getPassword() {
		return password;
	}

	public static void setPassword(char[] password) {
		ALMData.password = password;
	}

	public static String getDomain() {
		return domain;
	}

	public static void setDomain(String domain) {
		ALMData.domain = domain;
	}

	public static String getProject() {
		return project;
	}

	public static void setProject(String project) {
		ALMData.project = project;
	}

	public int getTestSetID() {
		return testSetID;
	}

	public void setTestSetID(int testSetID) {
		this.testSetID = testSetID;
	}

	public String getTestCaseName() {
		return testCaseName;
	}

	public void setTestCaseName(String testCaseName) {
		this.testCaseName = testCaseName;
	}

	public StatusAs getTestCaseStatus() {
		return testCaseStatus;
	}

	public void setTestCaseStatus(StatusAs testCaseStatus) {
		this.testCaseStatus = testCaseStatus;
	}

	public String getDefectID() {
		return defectID;
	}

	public void setDefectID(String defectID) {
		this.defectID = defectID;
	}

	public String getFailedTestStepNo() {
		return failedTestStepNo;
	}

	public void setFailedTestStepNo(String failedTestStepNo) {
		this.failedTestStepNo = failedTestStepNo;
	}

	public String getFailedTestStepActualResult() {
		return failedTestStepActualResult;
	}

	public void setFailedTestStepActualResult(String failedTestStepActualResult) {
		this.failedTestStepActualResult = failedTestStepActualResult;
	}

	public File getTestStepsXLFile() {
		return testStepsXLFile;
	}

	public void setTestStepsXLFile(File xlFile) {
		this.testStepsXLFile = xlFile;
	}

	public File getAttachmentFile() {
		return attachmentFile;
	}

	public void setAttachmentFile(File attachmentFile) {
		this.attachmentFile = attachmentFile;
	}
	
	public String getRunName() {
		return runName;
	}

	public void setRunName(String runName) {
		this.runName = runName;
	}

	public File getOutputXLFileLoc() {
		return outputXLFileLoc;
	}

	public void setOutputXLFileLoc(File outputXLFileLoc) {
		this.outputXLFileLoc = outputXLFileLoc;
	}

	public String getTestFolderPathOrTestSetID() {
		return testFolderPathOrTestSetID;
	}

	public void setTestFolderPathOrTestSetID(String testFolderPathOrTestSetID) {
		this.testFolderPathOrTestSetID = testFolderPathOrTestSetID;
	}
}
