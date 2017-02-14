package com.alm.wrapper.ui;

import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.IOException;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import com.alm.wrapper.classes.ALMAutomationWrapper;
import com.alm.wrapper.classes.ALMData;
import com.alm.wrapper.classes.FetchAttachmentDetails;
import com.alm.wrapper.enums.Domains;
import com.alm.wrapper.enums.Projects;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.LibraryLoader;

import jxl.write.WriteException;

/**
 * Class representing the ALM Login window of the tool. Contains main method to
 * trigger the execution.
 * 
 * @author sahil.srivastava
 *
 */
public class ALMLoginAndAttachmentWindow extends JFrame implements ActionListener {

	private static final long serialVersionUID = 1L;
	private static final String VERSION = "v2.0.0";

	private static final String FAILED_CONN_STRING = "Unable to connect to ALM...!!!";
	private static final String SUCCESS_CONN_STRING = "Connected to ALM...";
	private static final String AUTH_FAIL_STRING = "Authentication failed";
	private static final String INVALID_USERN_PASS_STRING = "Invalid username/password";
	private static final String EXCEPTION_OCCURED_STRING = "Exception Occured";

	private ALMAutomationWrapper almAutomationWrapper;
	private static ALMData almData;
	private FetchAttachmentDetails fetchDetails;

	private String almURLDefault = "ALM URL";

	private JPanel panel;

	// Declaring Labels
	private JLabel backgroundImgLabel;
	private JLabel almURLLabel;
	private JLabel userNameLabel;
	private JLabel passwordLabel;
	private JLabel domainLabel;
	private JLabel projectLabel;
	private JLabel logLabel;
	private JLabel testFolderPathOrTestSetIDLabel;
	private JLabel outputExcelPathLabel;
	private JLabel attachmentLogLabel;

	// Declaring Text Fields
	private JTextField almURLTextField;
	private JTextField userNameTextField;
	private JPasswordField passwordTextField;
	private JTextField testFolderPathOrTestSetIDField;
	private JTextField outputExcelPathField;
	
	private JButton outputLocXlButton;
	private JButton startFetchingButton;
	
	private JFileChooser fileChooser;

	private JComboBox<String> domainCombo;
	private String[] domainList = { Domains.DOMAIN_NAME1.getDomain(), Domains.DOMAIN_NAME2.getDomain() };

	private JComboBox<String> projectCombo;
	private String[] projectList = { Projects.PROJECT_NAME1.getProject(),
			Projects.PROJECT_NAME2.getProject() };

	// Declaring Buttons
	private JButton loginButton;
	private JButton clearButton;

	private boolean isConnected;

	public ALMLoginAndAttachmentWindow() {
		setSize(480, 300);
		setTitle("ALM Attachment Details Agent - " + VERSION);
		setTitle("ALM Attachment Details Fetcher");
		setLayout(null);

		setResizable(false);
		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		int w = getSize().width;
		int h = getSize().height;
		int x = (dim.width - w) / 2;
		int y = (dim.height - h) / 2;

		setLocation(x, y);

		addWindowListener(new java.awt.event.WindowAdapter() {
			@Override
			public void windowClosing(java.awt.event.WindowEvent windowEvent) {
				setVisible(false);
				if (almAutomationWrapper != null) {
					ActiveXComponent almActiveXComponent = almAutomationWrapper.getAlmActiveXComponent();
					almActiveXComponent.invoke("Disconnect");
					System.out.println("Disconnecting from project...");
					almActiveXComponent.invoke("Logout");
					System.out.println("Terminating the user's connection: logging out...");
					almActiveXComponent.invoke("ReleaseConnection");
					System.out.println("Releasing the COM pointer...");
				}
				System.exit(0);
			}
		});

		panel = new JPanel();
		panel.setLayout(null);
		setContentPane(panel);

		// Setting the logo for the frame
		ImageIcon imgIcon = new ImageIcon("");
		setIconImage(imgIcon.getImage());
		
		backgroundImgLabel = new JLabel(new ImageIcon(""));
		backgroundImgLabel.setBounds(0, -15, 480, 300);
		panel.add(backgroundImgLabel);
		
		// Adding almURLLabel
		almURLLabel = new JLabel("ALM URL");
		almURLLabel.setLabelFor(almURLTextField);
		almURLLabel.setBounds(50, 30, 100, 25);
		backgroundImgLabel.add(almURLLabel);

		// Adding almURLTextField
		almURLTextField = new JTextField(getAlmURLDefault(), 20);
		almURLTextField.setBounds(150, 30, 220, 23);
		almURLTextField.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(almURLTextField);

		// Adding userNameLabel
		userNameLabel = new JLabel("User Name");
		userNameLabel.setLabelFor(userNameTextField);
		userNameLabel.setBounds(50, 60, 100, 25);
		backgroundImgLabel.add(userNameLabel);

		// Adding userNameTextField
		userNameTextField = new JTextField(10);
		userNameTextField.setBounds(150, 60, 120, 23);
		userNameTextField.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(userNameTextField);

		// Adding passwordLabel
		passwordLabel = new JLabel("Password");
		passwordLabel.setLabelFor(passwordTextField);
		passwordLabel.setBounds(50, 90, 100, 25);
		backgroundImgLabel.add(passwordLabel);

		// Adding passwordTextField
		passwordTextField = new JPasswordField(10);
		passwordTextField.setBounds(150, 90, 120, 25);
		passwordTextField.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(passwordTextField);

		// Adding domainLabel
		domainLabel = new JLabel("Domain");
		domainLabel.setLabelFor(domainCombo);
		domainLabel.setBounds(50, 120, 100, 25);
		backgroundImgLabel.add(domainLabel);

		// Adding domainCombo
		domainCombo = new JComboBox<String>(domainList);
		domainCombo.setBounds(150, 120, 220, 25);
		domainCombo.setEditable(true);
		backgroundImgLabel.add(domainCombo);

		// Adding projectLabel
		projectLabel = new JLabel("Project");
		projectLabel.setLabelFor(projectCombo);
		projectLabel.setBounds(50, 150, 100, 25);
		backgroundImgLabel.add(projectLabel);

		// Adding projectCombo
		projectCombo = new JComboBox<String>(projectList);
		projectCombo.setBounds(150, 150, 220, 25);
		projectCombo.setEditable(true);
		backgroundImgLabel.add(projectCombo);

		// Adding log Label
		logLabel = new JLabel();
		logLabel.setBounds(150, 175, 200, 25);
		logLabel.setHorizontalAlignment((int) CENTER_ALIGNMENT);
		backgroundImgLabel.add(logLabel);

		// Adding loginButton
		loginButton = new JButton("Login");
		loginButton.setBounds(150, 200, 80, 25);
		loginButton.addActionListener(this);
		loginButton.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(loginButton);

		// Adding clearButton
		clearButton = new JButton("Clear");
		clearButton.setBounds(280, 200, 80, 25);
		clearButton.addActionListener(this);
		clearButton.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(clearButton);
		
		setVisible(true);
//		extendWindowAndTakeInputs();
	}

	public String getAlmURLDefault() {
		return almURLDefault;
	}

	public void setAlmURLDefault(String almURL) {
		this.almURLDefault = almURL;
	}

	@Override
	public void actionPerformed(ActionEvent event) {
		if (event.getSource() == loginButton) {
			logLabel.setText("Logging in...");
			logLabel.paintImmediately(logLabel.getVisibleRect());
			ALMData.setAlmURL(almURLTextField.getText());
			ALMData.setUserName(userNameTextField.getText());
			ALMData.setPassword(passwordTextField.getPassword());
			ALMData.setDomain(domainCombo.getSelectedItem().toString());
			ALMData.setProject(projectCombo.getSelectedItem().toString());

			almAutomationWrapper = new ALMAutomationWrapper(almData);
			System.out.println("ALM automation wrapper object created");

			try {
				isConnected = almAutomationWrapper.connectAndLoginALM();
				if (!isConnected) {
					logLabel.setText(FAILED_CONN_STRING);
				} else {
					System.out.println(SUCCESS_CONN_STRING);
					logLabel.setText(SUCCESS_CONN_STRING);
					extendWindowAndTakeInputs();
				}
			} catch (Exception e) {
				System.err.println(e.getMessage());
				if (e.getMessage().contains(AUTH_FAIL_STRING)) {
					logLabel.setText(INVALID_USERN_PASS_STRING);
				} else {
					logLabel.setText(FAILED_CONN_STRING);
				}
				almAutomationWrapper.closeConnection();
			}
		}

		else if (event.getSource() == clearButton) {
			userNameTextField.setText("");
			passwordTextField.setText("");
		}
		
		else if (event.getSource() == outputLocXlButton) {
			int returnVal = fileChooser.showOpenDialog(ALMLoginAndAttachmentWindow.this);
			if (returnVal == JFileChooser.APPROVE_OPTION) {
				almData.setOutputXLFileLoc(fileChooser.getSelectedFile());
				outputExcelPathField.setText(almData.getOutputXLFileLoc().getAbsolutePath());
			}
		}
		
		else if(event.getSource() == startFetchingButton) {
			almData.setTestFolderPathOrTestSetID(testFolderPathOrTestSetIDField.getText());
			fetchDetails = new FetchAttachmentDetails(almAutomationWrapper);
			try {
				fetchDetails.fetchAndOutputAttachmentDetails(this);
			} catch (WriteException | IOException e) {
				e.printStackTrace();
				updateAttachmentLogLabel(EXCEPTION_OCCURED_STRING);
			}
		}
	}

	public static void main(String... s) {
		System.setProperty("jacob.dll.path", System.getProperty("user.dir") + "\\jacob-1.18-x86.dll");
		LibraryLoader.loadJacobLibrary();

		// Changing the Look and Feel of the UI to Nimbus
//		try {
//			for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
//				if ("Nimbus".equals(info.getName())) {
//					UIManager.setLookAndFeel(info.getClassName());
//					break;
//				}
//			}
//		} catch (Exception e) {
//			try {
//				UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
//			} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
//					| UnsupportedLookAndFeelException e1) {
//				e1.printStackTrace();
//			}
//		}
		
		// Changing the Look And Feel of the UI to Native OS look and feel
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
				| UnsupportedLookAndFeelException e1) {
			e1.printStackTrace();
		}
		
		
		
		new ALMLoginAndAttachmentWindow();
		almData = new ALMData();
	}

	/**
	 * Implementation of KeyListener interface to invoke action events when
	 * specific keys are pressed
	 * 
	 * @author sahil.srivastava
	 *
	 */
	public class KeyListenerImplementation implements KeyListener {

		@Override
		public void keyTyped(KeyEvent arg0) {
			// Definition not required as of now
		}

		@Override
		public void keyReleased(KeyEvent arg0) {
			// Definition not required as of now
		}

		@Override
		public void keyPressed(KeyEvent event) {
			// If the focus is on clear button and enter key is pressed
			if (event.getSource() == clearButton && event.getKeyCode() == KeyEvent.VK_ENTER) {
				actionPerformed(new ActionEvent(clearButton, 2, ""));
			}
			// If enter key is pressed
			else if (event.getKeyCode() == KeyEvent.VK_ENTER) {
				actionPerformed(new ActionEvent(loginButton, 1, ""));
			}
		}
	}
	
	private void extendWindowAndTakeInputs() {
		setSize(550, 480);
		
		backgroundImgLabel = new JLabel(new ImageIcon(""));
		backgroundImgLabel.setBounds(0, 0, 550, 480);
		panel.add(backgroundImgLabel);
		
		// Adding testFolderPathLabel
		testFolderPathOrTestSetIDLabel = new JLabel("FolderPath Or TestSetID");
		testFolderPathOrTestSetIDLabel.setLabelFor(testFolderPathOrTestSetIDField);
		testFolderPathOrTestSetIDLabel.setBounds(50, 280, 150, 25);
		backgroundImgLabel.add(testFolderPathOrTestSetIDLabel);

		// Adding testFolderPathField
		testFolderPathOrTestSetIDField = new JTextField(20);
		testFolderPathOrTestSetIDField.setBounds(200, 280, 220, 23);
		testFolderPathOrTestSetIDField.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(testFolderPathOrTestSetIDField);
		
		// Adding outputExcelPathLabel
		outputExcelPathLabel = new JLabel("Output Excel Path");
		outputExcelPathLabel.setLabelFor(outputExcelPathField);
		outputExcelPathLabel.setBounds(50, 310, 100, 25);
		backgroundImgLabel.add(outputExcelPathLabel);
		
		// Adding outputExcelPathField
		outputExcelPathField = new JTextField(20);
		outputExcelPathField.setBounds(200, 310, 220, 23);
		outputExcelPathField.setEditable(false);
		outputExcelPathField.addKeyListener(new KeyListenerImplementation());
		backgroundImgLabel.add(outputExcelPathField);
		
		// Adding outputLocXlButton
		outputLocXlButton = new JButton("Browse");
		outputLocXlButton.setBounds(430, 310, 70, 23);
		outputLocXlButton.addActionListener(this);
		backgroundImgLabel.add(outputLocXlButton);
		fileChooser = new JFileChooser();
		fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		
		attachmentLogLabel = new JLabel();
		attachmentLogLabel.setBounds(180, 345, 200, 25);
		attachmentLogLabel.setHorizontalAlignment((int) LEFT_ALIGNMENT);
		backgroundImgLabel.add(attachmentLogLabel);
		
		// Adding startFetchingButton
		startFetchingButton = new JButton("Start fetching...!!!");
		startFetchingButton.setBounds(200, 380, 120, 25);
		startFetchingButton.setHorizontalAlignment((int) CENTER_ALIGNMENT);
		startFetchingButton.addActionListener(this);
		backgroundImgLabel.add(startFetchingButton);
		
		// Disabling all other UI components
		almURLTextField.setEnabled(false);
		userNameTextField.setEnabled(false);
		passwordTextField.setEnabled(false);
		
		projectCombo.setEnabled(false);
		domainCombo.setEnabled(false);
		
		loginButton.setEnabled(false);
		clearButton.setEnabled(false);
	}
	
	public void updateAttachmentLogLabel(String str) {
		attachmentLogLabel.setText(str);
		attachmentLogLabel.paintImmediately(logLabel.getVisibleRect());
	}
}