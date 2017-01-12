package com.alm.wrapper.classes;

import java.io.File;
import java.io.IOException;
import java.util.NoSuchElementException;

import com.alm.wrapper.ui.ALMLoginAndAttachmentWindow;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * 
 * @author sahil.srivastava
 *
 */
public class FetchAttachmentDetails {

	private ALMAutomationWrapper almAutomationWrapper;
	private ALMData almData;
	
	private final static int INITIAL_ATTACHMENT_NAME_COUNT = 6;
	private final static int INITIAL_ATTACHMENT_SIZE_COUNT = 7;
	
	private static int startColumnNo;
	
	private static int attachmentNameCount;
	private static int attachmentSizeCount;
	
	private static int maxAttachmentSizeColumnNo;
	
	private static Label testSetIDLabel;
	private static Label testCaseNameLabel;
	private static Label testCaseStatusLabel;
	private static Label hasAttachmentLabel;
	
	private static WritableWorkbook workbook;
	private static WritableSheet sheet;
	
	private int count = 3;
	
	private final static String FINISHED_STRING = "Finished fetching/writing details...!!!";
	
	private final static String ATTACHMENT_NAME_COMPLIANCE_CHECK_STRING = "Screenshot";
	
	
	public FetchAttachmentDetails(ALMAutomationWrapper almAutomationWrapper) {
		this.almAutomationWrapper = almAutomationWrapper;
		this.almData = almAutomationWrapper.getAlmData();
	}
	
	
	public void fetchAndOutputAttachmentDetails(ALMLoginAndAttachmentWindow window) throws RowsExceededException, WriteException, IOException  {
		
		boolean found = false;
		String[] testFolders;
		ActiveXComponent activeXComponent = null;
		ActiveXComponent testSet;
		workbook = createExcelFile();
		sheet = workbook.getSheet(0);
		
		try {

			window.updateAttachmentLogLabel("Fetching details...");
			activeXComponent = almAutomationWrapper.getAlmActiveXComponent();
			
			try {
				almData.setTestSetID(Integer.parseInt(almData.getTestFolderPathOrTestSetID()));
				testSet = activeXComponent.getPropertyAsComponent("TestSetFactory").invokeGetComponent("Item",
						new Variant(almData.getTestSetID()));
				fetchTestSetDataAndWriteInExcel(testSet, null);
				window.updateAttachmentLogLabel(FINISHED_STRING);
			} catch (NumberFormatException e) {
				ActiveXComponent testLabfolderFactory = activeXComponent.getPropertyAsComponent("TestLabFolderFactory");

				ActiveXComponent root = testLabfolderFactory.invokeGetComponent("Root");

				Variant testLabFoldersVariant = root.invoke("NewList");

				EnumVariant testFoldersEnumVariant = new EnumVariant(testLabFoldersVariant.getDispatch());
				
				ActiveXComponent testFolder = null;
				
				testFolders = almData.getTestFolderPathOrTestSetID().split("/");

				for (int i = 0; i < testFolders.length; i++) {
					found = false;
					while (testFoldersEnumVariant.hasMoreElements()) {
						testFolder = new ActiveXComponent(testFoldersEnumVariant.nextElement().getDispatch());
						if (testFolder.getPropertyAsString("Name").equals(testFolders[i])) {
							found = true;
							break;
						}
					}
					if (!found) {
						window.updateAttachmentLogLabel(testFolders[i] + " not found...");
						break;
					} else {
						testFoldersEnumVariant = new EnumVariant(testFolder.invoke("NewList").getDispatch());
					}
				}
				if (found) {
					fetchFolderDataAndWriteInExcel(testFolder);
					window.updateAttachmentLogLabel(FINISHED_STRING);
				}
			}
		} finally {
			workbook.write();
			workbook.close();
			
			System.out.println("Terminated...!!!");
		}
	}
	
	/**
	 * Recursive method to fetch all the test set/ test case data inside a
	 * folder. Would even fetch the data from the subfolders
	 * 
	 * @param testFolder
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	private void fetchFolderDataAndWriteInExcel(ActiveXComponent testFolder) throws RowsExceededException, WriteException {
		// Fetching the subfolders
		Variant subFoldersVariant = testFolder.invoke("NewList");
		try {
			EnumVariant subFoldersEnumVariant = new EnumVariant(subFoldersVariant.getDispatch());
			while (subFoldersEnumVariant.hasMoreElements()) {
				fetchFolderDataAndWriteInExcel(new ActiveXComponent(subFoldersEnumVariant.nextElement().getDispatch()));
			}
		} catch(Exception e) {
			// Fetching the test sets
			ActiveXComponent testSetFactory = testFolder.getPropertyAsComponent("TestSetFactory");
			Variant testSetVariant = testSetFactory.invoke("NewList", "");
			EnumVariant testSetEnumVariant = new EnumVariant(testSetVariant.getDispatch());

			while (testSetEnumVariant.hasMoreElements()) {
				fetchTestSetDataAndWriteInExcel(new ActiveXComponent(testSetEnumVariant.nextElement().getDispatch()), testFolder);
			}
		}

	}
	
	/**
	 * Method to fetch test cases from the test set and write the required details in excel 
	 * @param testSet
	 * @throws WriteException 
	 * @throws RowsExceededException 
	 */
	private void fetchTestSetDataAndWriteInExcel(ActiveXComponent testSet, ActiveXComponent testFolderOfTestSet) throws RowsExceededException, WriteException {
		int tempCount;
		float attachmentSize;
		boolean isAttachmentNameCompliant;
		String attachmentName;
		ActiveXComponent tsTest;
		
		Label attachmentNameLabel;
		Label attachmentSizeLabel;
		Label testParentFolderLabel;
		Label attachmentNameCompliantLabel;
		
		// If the execution is for a test set, testFolderOfTestSet will be passed as null
		if (testFolderOfTestSet == null) {
			testFolderOfTestSet = testSet.getPropertyAsComponent("TestSetFolder");
		}
		
		// Getting the testCaseFactory object pertaining to the testSet
		ActiveXComponent tsTestFactory = testSet.getPropertyAsComponent("TSTestFactory");
		
		// Fetching all the testCases in testCaseVariant
		Variant testCasesVariant = tsTestFactory.invoke("NewList", "");
		EnumVariant enumVariant = new EnumVariant(testCasesVariant.getDispatch());
		while(enumVariant.hasMoreElements()) {
			
			attachmentNameCount = INITIAL_ATTACHMENT_NAME_COUNT;
			attachmentSizeCount = INITIAL_ATTACHMENT_SIZE_COUNT;
			
			isAttachmentNameCompliant = true;
			
			testParentFolderLabel = new Label(startColumnNo, count, testFolderOfTestSet.getPropertyAsString("Name"), createFormatCellStatus(false));
			sheet.addCell(testParentFolderLabel);
			
			//Picking each test in the test set
			tsTest = new ActiveXComponent(enumVariant.nextElement().getDispatch());
			
			//Adding cell for test set id in excel sheet
			testSetIDLabel = new Label(startColumnNo + 1, count, testSet.getProperty("ID").toString(), createFormatCellStatus(false));
			sheet.addCell(testSetIDLabel);
			
			//Adding cell for test case name in excel sheet
			testCaseNameLabel = new Label(startColumnNo + 2, count, tsTest.getPropertyAsString("TestName"), createFormatCellStatus(false));
			sheet.addCell(testCaseNameLabel);
			System.out.println(tsTest.getPropertyAsString("TestName") + "   ");
			
			
			// Adding cell for test case status in excel sheet
			testCaseStatusLabel = new Label(startColumnNo + 3, count, tsTest.getPropertyAsString("Status"), createFormatCellStatus(false));
			sheet.addCell(testCaseStatusLabel);
			
			// Fetching the attachment factory from the test case
			ActiveXComponent attachmentFactory = tsTest.getPropertyAsComponent("Attachments");
			Variant variantAttachment = new Variant("");
			
			//Fetching the attachment list from the attachment factory
			Variant attachmentsVariant = attachmentFactory.invoke("NewList", variantAttachment);
			EnumVariant enumAttachVariant = new EnumVariant(attachmentsVariant.getDispatch());
			ActiveXComponent attachment;
			
			try {
				do {
					if(maxAttachmentSizeColumnNo < attachmentSizeCount) {
						maxAttachmentSizeColumnNo = attachmentSizeCount;
					}
					attachment = new ActiveXComponent(enumAttachVariant.nextElement().getDispatch());

					// Adding cell with attachment name as content
					attachmentName = attachment.getProperty("Name").toString();
					if(attachmentName.startsWith("TESTCYCL_")) {
						attachmentName = attachmentName.substring(16);
					}
					
					// Logic to check attachment name compliance
					if (isAttachmentNameCompliant) {
						try {
							if (!attachmentName.split("_")[3].equalsIgnoreCase(ATTACHMENT_NAME_COMPLIANCE_CHECK_STRING)) {
								isAttachmentNameCompliant = false;
							}
						} catch (ArrayIndexOutOfBoundsException arrayException) {
							isAttachmentNameCompliant = false;
						}

					}
					
					attachmentNameLabel = new Label(attachmentNameCount, count, attachmentName, createFormatCellStatus(false));
					sheet.addCell(attachmentNameLabel);
					attachmentNameCount += 2;

					// Adding cell with attachment size as content
					attachmentSize = Float.parseFloat(attachment.getProperty("FileSize").toString());
					attachmentSize /= 1024;
					
					attachmentSizeLabel = new Label(attachmentSizeCount, count, Math.round(attachmentSize * 100)/100f + "", createFormatCellStatus(false));
					sheet.addCell(attachmentSizeLabel);
					attachmentSizeCount += 2;
				} while (enumAttachVariant.hasMoreElements());
				// Adding cell with value 'Yes' as attachment is present for
				// the test case
				hasAttachmentLabel = new Label(startColumnNo + 4, count, "Yes", createFormatCellStatus(false));
				sheet.addCell(hasAttachmentLabel);
				
				// Adding cell with value 'Yes/No' for attachment compliance
				if(isAttachmentNameCompliant) {
				attachmentNameCompliantLabel = new Label(startColumnNo + 5, count, "Yes", createFormatCellStatus(false));
				sheet.addCell(attachmentNameCompliantLabel);
				}
				else {
					attachmentNameCompliantLabel = new Label(startColumnNo + 5, count, "No", createFormatCellStatus(false));
					sheet.addCell(attachmentNameCompliantLabel);
				}
			}
			// NoSuchElementException is thrown if there is no attachment
			// present in the list
			catch (NoSuchElementException exp) {
				// Adding cell with value 'No' as no attachment is there for
				// the test case and proceeding with the next test case
				hasAttachmentLabel = new Label(startColumnNo + 4, count, "No", createFormatCellStatus(false));
				sheet.addCell(hasAttachmentLabel);
				
				// Adding cell with value 'NA' for attachment compliance as no attachment is there
				attachmentNameCompliantLabel = new Label(startColumnNo + 5, count, "NA", createFormatCellStatus(false));
				sheet.addCell(attachmentNameCompliantLabel);
				
				count++;
				continue;
			}
			count++;
		}
		tempCount = INITIAL_ATTACHMENT_SIZE_COUNT;
		while(tempCount < maxAttachmentSizeColumnNo) {
			tempCount += 2;
			sheet.addCell(new Label(tempCount-1, 2, "Attachment Name", createFormatCellStatus(true)));
			sheet.addCell(new Label(tempCount, 2, "Attachment Size (In KB)", createFormatCellStatus(true)));
		}
	}
	
	
	
	/**
	 * Method to create excel file
	 * 
	 * @return
	 * @throws IOException
	 * @throws RowsExceededException
	 * @throws WriteException
	 */
	private WritableWorkbook createExcelFile() throws IOException, RowsExceededException, WriteException {
		WritableWorkbook workbook = Workbook.createWorkbook(new File(almData.getOutputXLFileLoc().getAbsolutePath() + "\\attachmentDetails.xls"));
		WritableSheet sheet = workbook.createSheet("AttachmentDetails", 0);
		
		Label testParentFolder = new Label(startColumnNo, 2, "Test Parent Folder", createFormatCellStatus(true)); //column=0=A,row=0=1
		sheet.addCell(testParentFolder);
		
		Label testSetIDLabel = new Label(startColumnNo + 1, 2, "TestSet ID", createFormatCellStatus(true)); //column=0=A,row=0=1
		sheet.addCell(testSetIDLabel);
		
		Label testCaseNameLabel = new Label(startColumnNo + 2, 2, "Test Case Name", createFormatCellStatus(true));
		sheet.addCell(testCaseNameLabel);
		
		Label testCaseStatusLabel = new Label(startColumnNo + 3, 2, "Test Case Status", createFormatCellStatus(true));
		sheet.addCell(testCaseStatusLabel);
		
		Label hasAttachmentLabel = new Label(startColumnNo + 4, 2, "Has Attachment", createFormatCellStatus(true));
		sheet.addCell(hasAttachmentLabel);
		
		Label attachmentNameCompliant = new Label(startColumnNo + 5, 2, "Attachment Name Compliant", createFormatCellStatus(true));
		sheet.addCell(attachmentNameCompliant);
		
		Label attachmentName = new Label(INITIAL_ATTACHMENT_NAME_COUNT, 2, "Attachment Name", createFormatCellStatus(true));
		sheet.addCell(attachmentName);
		
		Label attachmentSizeLabel = new Label(INITIAL_ATTACHMENT_SIZE_COUNT, 2, "Attachment Size (In KB)", createFormatCellStatus(true));
		sheet.addCell(attachmentSizeLabel);
		
		return workbook;
	}
	
	// Method to return WritableCellFormat instance after setting ceratin
	// properties for the cell
	public WritableCellFormat createFormatCellStatus(boolean isHeading) throws WriteException {
		WritableFont wfontStatus;
		WritableCellFormat fCellstatus;
		if (isHeading) {
			wfontStatus = new WritableFont(WritableFont.createFont("Arial"), WritableFont.DEFAULT_POINT_SIZE,
					WritableFont.BOLD, false, UnderlineStyle.NO_UNDERLINE);
		} else {
			wfontStatus = new WritableFont(WritableFont.createFont("Arial"), WritableFont.DEFAULT_POINT_SIZE,
					WritableFont.NO_BOLD, false, UnderlineStyle.NO_UNDERLINE);
		}
		fCellstatus = new WritableCellFormat(wfontStatus);
		fCellstatus.setAlignment(jxl.format.Alignment.CENTRE);
		fCellstatus.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);

		if (isHeading) {
			fCellstatus.setBackground(Colour.GRAY_25);
			fCellstatus.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.MEDIUM);
			fCellstatus.setWrap(true);
		}

		return fCellstatus;
	}
}