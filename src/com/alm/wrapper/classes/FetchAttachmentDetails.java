package com.alm.wrapper.classes;

import java.io.File;
import java.io.IOException;
import java.util.NoSuchElementException;

import com.alm.wrapper.enums.ExecutionStatus;
import com.alm.wrapper.ui.ALMLoginAndAttachmentWindow;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComFailException;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
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
	private ALMLoginAndAttachmentWindow window;

	private final static int INITIAL_ATTACHMENT_NAME_COUNT = 7;
	private final static int INITIAL_ATTACHMENT_SIZE_COUNT = 8;

	private static int startColumnNo;

	private static int attachmentNameCount;
	private static int attachmentSizeCount;

	private static int maxAttachmentSizeColumnNo;

	private static Label testParentFolderLabel;
	private static Label testSetIDLabel;
	private static Label testCaseNameLabel;
	private static Label testIDLabel;
	private static Label testerLabel;
	private static Label hasAttachmentLabel;
	private static Label attachmentNameCompliantLabel;

	private static WritableWorkbook workbook;
	private static WritableSheet sheet;

	private int count = 3;

	private final static String FINISHED_STRING = "Finished fetching/writing details...!!!";

	private final static String MANDATED_SCREENSHOT_TEXT_CONVENTION = "Screenshot";
	private final static String MANDATED_LOG_TEXT_CONVENTION = "Logs";

	private boolean exceptionOccured;

	public FetchAttachmentDetails(ALMAutomationWrapper almAutomationWrapper) {
		this.almAutomationWrapper = almAutomationWrapper;
		this.almData = almAutomationWrapper.getAlmData();
	}

	public void fetchAndOutputAttachmentDetails(ALMLoginAndAttachmentWindow window)
			throws RowsExceededException, WriteException, IOException {

		boolean found = false;
		String[] testFolders;
		ActiveXComponent activeXComponent = null;
		ActiveXComponent testSet;
		this.window = window;
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
				if (!exceptionOccured) {
					window.updateAttachmentLogLabel(FINISHED_STRING);
				}
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
					if (!exceptionOccured) {
						window.updateAttachmentLogLabel(FINISHED_STRING);
					}
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
	private void fetchFolderDataAndWriteInExcel(ActiveXComponent testFolder)
			throws RowsExceededException, WriteException {
		// Fetching the subfolders
		Variant subFoldersVariant = testFolder.invoke("NewList");
		try {
			EnumVariant subFoldersEnumVariant = new EnumVariant(subFoldersVariant.getDispatch());
			while (subFoldersEnumVariant.hasMoreElements()) {
				fetchFolderDataAndWriteInExcel(new ActiveXComponent(subFoldersEnumVariant.nextElement().getDispatch()));
			}
		} catch (Exception e) {
			// Fetching the test sets
			ActiveXComponent testSetFactory = testFolder.getPropertyAsComponent("TestSetFactory");
			Variant testSetVariant = testSetFactory.invoke("NewList", "");
			EnumVariant testSetEnumVariant = new EnumVariant(testSetVariant.getDispatch());

			while (testSetEnumVariant.hasMoreElements()) {
				fetchTestSetDataAndWriteInExcel(new ActiveXComponent(testSetEnumVariant.nextElement().getDispatch()),
						testFolder);
			}
		}

	}

	/**
	 * Method to fetch test cases from the test set and write the required
	 * details in excel
	 * 
	 * @param testSet
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	private void fetchTestSetDataAndWriteInExcel(ActiveXComponent testSet, ActiveXComponent testFolderOfTestSet)
			throws RowsExceededException, WriteException {
		int tempCount;
		float attachmentSize;
		boolean isAttachmentNameCompliant;
		String attachmentName;
		String testID;
		ActiveXComponent tsTest;

		Label attachmentNameLabel;
		Label attachmentSizeLabel;

		try {
			// If the execution is for a test set, testFolderOfTestSet will be
			// passed as null
			if (testFolderOfTestSet == null) {
				testFolderOfTestSet = testSet.getPropertyAsComponent("TestSetFolder");
			}

			// Getting the testCaseFactory object pertaining to the testSet
			ActiveXComponent tsTestFactory = testSet.getPropertyAsComponent("TSTestFactory");

			// Fetching all the testCases in testCaseVariant
			Variant testCasesVariant = tsTestFactory.invoke("NewList", "");
			EnumVariant enumVariant = new EnumVariant(testCasesVariant.getDispatch());
			while (enumVariant.hasMoreElements()) {

				attachmentNameCount = INITIAL_ATTACHMENT_NAME_COUNT;
				attachmentSizeCount = INITIAL_ATTACHMENT_SIZE_COUNT;

				isAttachmentNameCompliant = true;

				testParentFolderLabel = new Label(startColumnNo, count, testFolderOfTestSet.getPropertyAsString("Name"),
						createFormatCellStatus(false));
				sheet.addCell(testParentFolderLabel);

				// Picking each test in the test set
				tsTest = new ActiveXComponent(enumVariant.nextElement().getDispatch());

				// Adding cell for test case status in excel sheet
				if (tsTest.getPropertyAsString("Status").equals(ExecutionStatus.PASSED.getStatus())) {

					// Adding cell for test set id in excel sheet
					testSetIDLabel = new Label(startColumnNo + 1, count, testSet.getProperty("ID").toString(),
							createFormatCellStatus(false));
					sheet.addCell(testSetIDLabel);

					// Adding cell for test case name in excel sheet
					testCaseNameLabel = new Label(startColumnNo + 2, count, tsTest.getPropertyAsString("TestName"),
							createFormatCellStatus(false));
					sheet.addCell(testCaseNameLabel);
					System.out.println(tsTest.getPropertyAsString("TestName") + "   ");

					// Adding cell for test id in excel sheet
					testID = tsTest.getProperty("TestId").toString();
					testIDLabel = new Label(startColumnNo + 3, count, testID, createFormatCellStatus(false));
					sheet.addCell(testIDLabel);

					// Adding cell for tester in excel sheet
					testerLabel = new Label(startColumnNo + 4, count,
							tsTest.invoke("Field", new Variant("TC_ACTUAL_TESTER")).toString(),
							createFormatCellStatus(false));
					sheet.addCell(testerLabel);

					// Fetching the attachment factory from the test case
					ActiveXComponent attachmentFactory = tsTest.getPropertyAsComponent("Attachments");
					Variant variantAttachment = new Variant("");

					// Fetching the attachment list from the attachment factory
					Variant attachmentsVariant = attachmentFactory.invoke("NewList", variantAttachment);
					EnumVariant enumAttachVariant = new EnumVariant(attachmentsVariant.getDispatch());
					ActiveXComponent attachment;

					try {
						do {
							if (maxAttachmentSizeColumnNo < attachmentSizeCount) {
								maxAttachmentSizeColumnNo = attachmentSizeCount;
							}
							attachment = new ActiveXComponent(enumAttachVariant.nextElement().getDispatch());

							// Adding cell with attachment name as content
							attachmentName = attachment.getProperty("Name").toString();
							if (attachmentName.startsWith("TESTCYCL_")) {
								attachmentName = attachmentName.substring(16);
							}
							// Validate attachment name compliance
							if (isAttachmentNameCompliant) {
								isAttachmentNameCompliant = validateAttachmentName(attachmentName, testID);
							}

							attachmentNameLabel = new Label(attachmentNameCount, count, attachmentName,
									createFormatCellStatus(false));
							sheet.addCell(attachmentNameLabel);
							attachmentNameCount += 2;

							// Adding cell with attachment size as content
							attachmentSize = Float.parseFloat(attachment.getProperty("FileSize").toString());
							attachmentSize /= 1024;

							attachmentSizeLabel = new Label(attachmentSizeCount, count,
									Math.round(attachmentSize * 100) / 100f + "", createFormatCellStatus(false));
							sheet.addCell(attachmentSizeLabel);
							attachmentSizeCount += 2;
						} while (enumAttachVariant.hasMoreElements());
						// Adding cell with value 'Yes' as attachment is present
						// for
						// the test case
						hasAttachmentLabel = new Label(startColumnNo + 5, count, "Yes", createFormatCellStatus(false));
						sheet.addCell(hasAttachmentLabel);

						// Adding cell with value 'Yes/No' for attachment
						// compliance
						if (isAttachmentNameCompliant) {
							attachmentNameCompliantLabel = new Label(startColumnNo + 6, count, "Yes",
									createFormatCellStatus(false));
							sheet.addCell(attachmentNameCompliantLabel);
						} else {
							attachmentNameCompliantLabel = new Label(startColumnNo + 6, count, "No",
									createFormatCellStatus(false));
							sheet.addCell(attachmentNameCompliantLabel);
						}
					}
					// NoSuchElementException is thrown if there is no
					// attachment
					// present in the list
					catch (NoSuchElementException exp) {
						// Adding cell with value 'No' as no attachment is there
						// for
						// the test case and proceeding with the next test case
						hasAttachmentLabel = new Label(startColumnNo + 5, count, "No", createFormatCellStatus(false));
						sheet.addCell(hasAttachmentLabel);

						// Adding cell with value 'NA' for attachment compliance
						// as
						// no attachment is there
						attachmentNameCompliantLabel = new Label(startColumnNo + 6, count, "NA",
								createFormatCellStatus(false));
						sheet.addCell(attachmentNameCompliantLabel);

						count++;
						continue;
					}
					count++;
				}
				tempCount = INITIAL_ATTACHMENT_SIZE_COUNT;
				while (tempCount < maxAttachmentSizeColumnNo) {
					tempCount += 2;
					sheet.addCell(new Label(tempCount - 1, 2,
							"Attachment" + (tempCount - INITIAL_ATTACHMENT_SIZE_COUNT) + " Name",
							createFormatCellStatus(true)));
					sheet.addCell(new Label(tempCount, 2,
							"Attachment" + (tempCount - INITIAL_ATTACHMENT_SIZE_COUNT) + " Size (In KB)",
							createFormatCellStatus(true)));
				}
			}
		} catch (ComFailException cfe) {
			if (cfe.getLocalizedMessage().contains("does not exist in table 'CYCLE'")) {
				exceptionOccured = true;
				window.updateAttachmentLogLabel("Test Set ID: " + almData.getTestSetID() + " does not exist...!!!");
			}
			cfe.printStackTrace();
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
		WritableWorkbook workbook = Workbook
				.createWorkbook(new File(almData.getOutputXLFileLoc().getAbsolutePath() + "\\attachmentDetails.xls"));
		WritableSheet sheet = workbook.createSheet("AttachmentDetails", 0);

		Label testParentFolder = new Label(startColumnNo, 2, "Test Parent Folder", createFormatCellStatus(true)); // column=0=A,row=0=1
		sheet.addCell(testParentFolder);

		Label testSetIDLabel = new Label(startColumnNo + 1, 2, "TestSet ID", createFormatCellStatus(true)); // column=0=A,row=0=1
		sheet.addCell(testSetIDLabel);

		Label testCaseNameLabel = new Label(startColumnNo + 2, 2, "Test Case Name", createFormatCellStatus(true));
		sheet.addCell(testCaseNameLabel);

		Label testIDLabel = new Label(startColumnNo + 3, 2, "Test ID", createFormatCellStatus(true));
		sheet.addCell(testIDLabel);

		Label testerLabel = new Label(startColumnNo + 4, 2, "Tester", createFormatCellStatus(true));
		sheet.addCell(testerLabel);

		Label hasAttachmentLabel = new Label(startColumnNo + 5, 2, "Has Attachment", createFormatCellStatus(true));
		sheet.addCell(hasAttachmentLabel);

		Label attachmentNameCompliant = new Label(startColumnNo + 6, 2, "Attachment Name Compliant",
				createFormatCellStatus(true));
		sheet.addCell(attachmentNameCompliant);

		Label attachmentName = new Label(INITIAL_ATTACHMENT_NAME_COUNT, 2, "Attachment1 Name",
				createFormatCellStatus(true));
		sheet.addCell(attachmentName);

		Label attachmentSizeLabel = new Label(INITIAL_ATTACHMENT_SIZE_COUNT, 2, "Attachment1 Size (In KB)",
				createFormatCellStatus(true));
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
		fCellstatus.setAlignment(Alignment.CENTRE);
		fCellstatus.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);

		if (isHeading) {
			fCellstatus.setBackground(Colour.GRAY_25);
			fCellstatus.setBorder(Border.ALL, BorderLineStyle.MEDIUM);
			fCellstatus.setWrap(true);
		}
		return fCellstatus;
	}

	private boolean validateAttachmentName(String attachmentName, String testID) {
		String attachmentNameWithoutExtension;
		String attachmentExtension;
		boolean isAttachmentNameCompliant = true;
		String[] strUnderscoreDelimiter;
		String[] strDotDelimiter;

		strDotDelimiter = attachmentName.split("\\.");
		attachmentNameWithoutExtension = strDotDelimiter[0];
		strUnderscoreDelimiter = attachmentNameWithoutExtension.split("_");
		if (strDotDelimiter.length == 2) {
			attachmentExtension = strDotDelimiter[1];
			try {
				if (attachmentExtension.equals("log") || attachmentExtension.equals("txt")) {
					if (!strUnderscoreDelimiter[2].equalsIgnoreCase(MANDATED_LOG_TEXT_CONVENTION)) {
						return false;
					}
				} else {
					if (!strUnderscoreDelimiter[2].equalsIgnoreCase(MANDATED_SCREENSHOT_TEXT_CONVENTION)) {
						return false;
					}
				}
			} catch (ArrayIndexOutOfBoundsException arrayException) {
				return false;
			}
		}
		else{
			if(!strUnderscoreDelimiter[2].equalsIgnoreCase(MANDATED_LOG_TEXT_CONVENTION) && !strUnderscoreDelimiter[2].equalsIgnoreCase(MANDATED_SCREENSHOT_TEXT_CONVENTION)) {
				return false;
			}
		}
		if (!strUnderscoreDelimiter[0].equals(testID)) {
			return false;
		}
		return isAttachmentNameCompliant;
	}
}