/**This is a simple program made for my internship. Due to the fact that I never signed a NDA, or any sort of contract really, I think it's in my right to 
put this program on github. 
The project essentially takes in user inputs and creates and excel sheet based on the parameters submitted.
Afterward, it will send an email of the finished file. The program is mostly bugtested, so I feel comfortable putting it on my github.

The one thing that has been changed is some code. All filepaths, email addresses, and passwords have been deleted. If you have any interest in using this program,
be sure that you fix those filepaths to accomodate your computer. That aside, this project is pretty open to anyone.**/

import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.AbstractButton;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JTextField;
import javax.swing.JScrollPane;
import java.util.*;
import java.io.*;
import java.nio.file.DirectoryNotEmptyException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.activation.*;
import org.apache.poi.hssf.usermodel.*;



public class ButtonDemoPanel2 extends javax.swing.JPanel {
	
	private static final long serialVersionUID = 1L;
	private Path path;
	String Submission, filename, extraparams = "  ", extrafields = "  "; 
	JScrollPane Scroller;
	private JButton Submit, AddExtra, OpenExtra;
	private JTextField FirstName, LastName, MiddleName, Company, PhoneExtension, EmailGroup, 
	DomainGroup, EvoExtra, Kayako, Jira, Conf, 
	Svn, BitBucket, Bamboo, Office, Dfa, LoginCred, ExtraField, ExtraParam;
	int number = 99;
	private String FirstNameS, LastNameS, MiddleNameS, EmailS, PhoneExtensionS, EmailGroupS, DomainGroupS, EvoExtraS, AccessLevelS, ClientExtraS, KayakoS, JiraS, ConfS, SvnS, BitBucketS, BambooS, 
	OfficeS, DfaS, LoginCredS;
	private JRadioButton ConfirmCompanyYes, ConfirmCompanyNo, 
	PhoneExtensionY, PhoneExtensionN, EmailGY, EmailGN, DomainADY, DomainADN,
	EvoEAY, EvoEAN, ClientEY, ClientEN, JiraY, JiraN, ConfY, ConfN, SvnY, SvnN,
	BitBY, BitBN, UAEM, USEM, KeepFile, DeleteFile;
	private ButtonGroup CompanyGroup, PhoneExtensionGroup, EmailGGroup, DomainGroupGroup,
	EvoEAGroup, ClientEGroup, JiraGroup, ConfGroup, SvnGroup, BitGroup, EmailFormatGroup, SaveFileGroup;
	private JLabel TextName, EmailT, PhoneTextYNT, EmailGYNT,DomainGroupYNT, 
	EvoExtraYNT, AccessLevelT, 
	ClientExtraYNT, KayakoT, JiraGT, ConfGT, SvnYNT, BitBucketYNT, BambooT, OfficeT, DfaT, LoginCredT, ConfirmSubmission, ConfirmEmail, BambooRole,
	KeepFileT;
	private ArrayList<String> require, info;
	private boolean PhoneExtensionB = false, EmailGB = false, DomainADB = false, JiraB = false, ConfB = false, SvnB = false, BitB = false, EmailUS, KeepEmailB = false, SendEmail = false, extravisible = false;
	String[] Roles = new String[] {"---", "Manager", "Employee"}, Numbers;
	String[] AccessLevelList = new String[] {"---", "Manager Level", "Employee Level"};
	String[] ClientExtraLists = new String[] {"---", "Yes", "No"};
	ArrayList<String> numbers, extraparamstring;
	JComboBox<String> NumberList, RoleList, AccessList, ClientExtraList;


	public ButtonDemoPanel2() {


		//JPanel per each line is here
		//Some JPanel's aren't used. They're extra in case things go south later on.
		JPanel OverLine = new JPanel();
		JPanel Line2 = new JPanel();
		JPanel Line2P1 = new JPanel();
		JPanel Line2P2 = new JPanel();
		JPanel Line3 = new JPanel();
		JPanel Line3P1 = new JPanel();
		JPanel Line3P2 = new JPanel();
		JPanel Line4 = new JPanel();
		JPanel Line5 = new JPanel();
		JPanel Line5P1 = new JPanel();
		JPanel Line5P2 = new JPanel();
		JPanel Line6 = new JPanel();
		JPanel Line7 = new JPanel();
		JPanel Line7P1 = new JPanel();
		JPanel Line7P2 = new JPanel();
		JPanel Line8 = new JPanel();
		JPanel Line8P1 = new JPanel();
		JPanel Line8P2 = new JPanel();
		JPanel Line9 = new JPanel();
		JPanel Line10 = new JPanel();
		JPanel Line10P1 = new JPanel();
		JPanel Line10P2 = new JPanel();
		JPanel Line11 = new JPanel();
		JPanel Line12 = new JPanel();
		JPanel Line12P1 = new JPanel();
		JPanel Line12P2 = new JPanel();
		JPanel Line13 = new JPanel();
		JPanel Line13P1 = new JPanel();
		JPanel Line13P2 = new JPanel();
		JPanel Line14 = new JPanel();
		JPanel Line14P1 = new JPanel();
		JPanel Line14P2 = new JPanel();
		JPanel Line15 = new JPanel();
		JPanel Line15P1 = new JPanel();
		JPanel Line15P2 = new JPanel();
		JPanel Line16 = new JPanel();
		JPanel Line17 = new JPanel();
		JPanel Line18 = new JPanel();
		JPanel Line19 = new JPanel();
		JPanel Line20 = new JPanel();
		JPanel Line21 = new JPanel();
		JPanel Line22 = new JPanel();
		JPanel Line22P1 = new JPanel();
		JPanel Line22P2 = new JPanel();
		JPanel Line23 = new JPanel();
		JPanel Line24 = new JPanel();
		JPanel Line25 = new JPanel();
		JPanel Line26 = new JPanel();
		JPanel Line27 = new JPanel();
		
		//grid layouts are here
		//Line's are formatted into individual lines on the submission program 
		//Any line with a P1 and P2 are used to better format the Panel. Keeps everything looking on the same level
		OverLine.setLayout(new GridLayout(28, 0));
		Line2.setLayout(new GridLayout(1, 2));
		Line2P1.setLayout(new GridLayout(1, 2));
		Line2P2.setLayout(new GridLayout(1, 2));
		Line3.setLayout(new GridLayout(1, 4));
		Line3P1.setLayout(new GridLayout(1, 1));
		Line3P2.setLayout(new GridLayout(1, 3));
		Line4.setLayout(new GridLayout(1, 4));
		Line5.setLayout(new GridLayout(1, 2));
		Line5P1.setLayout(new GridLayout(1, 1));
		Line5P2.setLayout(new GridLayout(1, 3));
		Line6.setLayout(new GridLayout(1, 4));
		Line7.setLayout(new GridLayout(1, 4));
		Line7P1.setLayout(new GridLayout(1, 1));
		Line7P2.setLayout(new GridLayout(1, 3));
		Line8.setLayout(new GridLayout(1, 4));
		Line8P1.setLayout(new GridLayout(1, 1));
		Line8P2.setLayout(new GridLayout(1, 3));
		Line9.setLayout(new GridLayout(1, 4));
		Line10.setLayout(new GridLayout(1, 4));
		Line10P1.setLayout(new GridLayout(1, 1));
		Line10P2.setLayout(new GridLayout(1, 3));
		Line11.setLayout(new GridLayout(1, 4));
		Line12.setLayout(new GridLayout(1, 4));
		Line12P1.setLayout(new GridLayout(1, 1));
		Line12P2.setLayout(new GridLayout(1, 3));
		Line13.setLayout(new GridLayout(1, 4));
		Line13P1.setLayout(new GridLayout(1, 1));
		Line13P2.setLayout(new GridLayout(1, 3));
		Line14.setLayout(new GridLayout(1, 4));
		Line14P1.setLayout(new GridLayout(1, 1));
		Line14P2.setLayout(new GridLayout(1, 3));
		Line15.setLayout(new GridLayout(1, 4));
		Line15P1.setLayout(new GridLayout(1, 1));
		Line15P2.setLayout(new GridLayout(1, 3));
		Line16.setLayout(new GridLayout(1, 4));
		Line17.setLayout(new GridLayout(1, 4));
		Line18.setLayout(new GridLayout(1, 4));
		Line19.setLayout(new GridLayout(1, 4));
		Line20.setLayout(new GridLayout(1, 4));
		Line21.setLayout(new GridLayout(1, 4));
		Line22.setLayout(new GridLayout(1, 4));
		Line22P1.setLayout(new GridLayout(1, 1));
		Line22P2.setLayout(new GridLayout(1, 3));
		Line23.setLayout(new GridLayout(1, 4));
		Line24.setLayout(new GridLayout(1, 2));
		Line25.setLayout(new GridLayout(1, 2));
		Line26.setLayout(new GridLayout(1, 2));
		Line27.setLayout(new GridLayout(1, 2));
		//This loop is a placeholder until I find out what specific number fields are being used
		numbers = new ArrayList<String>();
		numbers.add("---");
		for(int i = 0; i < 300; i++) {
			number++;			
			numbers.add(String.valueOf(number));
		}
		Numbers = numbers.toArray(new String[numbers.size()]);
				
			
		Submission = ""; //String of submissions
		
		// --All  JLabels are Inserted Here--
		TextName = new JLabel("Insert Full Name (First, Last, Middle)*");	
		EmailT = new JLabel("Email Format*");
		new JLabel("Company");
		new JLabel("Write Company");
		PhoneTextYNT = new JLabel("Phone Extension");		
		new JLabel("Phone Extension"); 
		EmailGYNT = new JLabel("Email Group");
		new JLabel("Email Group");
		DomainGroupYNT = new JLabel("Domain AD Groups");
		new JLabel("Domain AD Group");
		EvoExtraYNT = new JLabel("Evogence Extranet Access");
		new JLabel("Evogence Extranet Access");
		AccessLevelT = new JLabel("Access Level *");
		ClientExtraYNT = new JLabel("Confirm Client Extranet");
		new JLabel("Client Extranet");
		KayakoT = new JLabel("Kayako Account*");
		JiraGT = new JLabel("Jira Account");
		new JLabel("JIRA");
		ConfGT = new JLabel("Confluence");
		new JLabel("CONFLUENCE");
		SvnYNT = new JLabel("SVN");
		new JLabel("SVN");
		BitBucketYNT = new JLabel("BitBucket");
		new JLabel("BitBucket");
		BambooT = new JLabel("Bamboo *");
		OfficeT = new JLabel("MSOffice *");
		DfaT = new JLabel("Directory File Access *");
		LoginCredT = new JLabel("Log In Credentials *");
		BambooRole = new JLabel("Bamboo Role");
		KeepFileT = new JLabel("Keep File After Finished?");		

		//-------------Text fields for user entry are here--------------------------
		//-------------Scrollbar entry fields are also entered here-----------------
				
		FirstName = new JTextField();
		LastName = new JTextField();
		MiddleName = new JTextField();				
		FirstName.setText("First Name");
		LastName.setText("Last Name");
		MiddleName.setText("Middle Initial");	       			  		  								
		Company = new JTextField();
		Company.setEditable(false);		
		NumberList = new JComboBox<>(Numbers);
		NumberList.setEditable(true);
		EmailGroup = new JTextField();
		EmailGroup.setEditable(false);
		DomainGroup = new JTextField();
		DomainGroup.setEditable(false);
		EvoExtra = new JTextField();
		EvoExtra.setEditable(false);
		AccessList = new JComboBox<>(AccessLevelList);
		AccessList.setEditable(true);
		ClientExtraList = new JComboBox<>(ClientExtraLists);
		Kayako = new JTextField();
		Jira = new JTextField();
		Jira.setEditable(false);
		Conf = new JTextField();
		Conf.setEditable(false);
		Svn = new JTextField();
		Svn.setEditable(false);
		BitBucket = new JTextField();
		BitBucket.setEditable(false);
		Bamboo = new JTextField(); 
		RoleList = new JComboBox<>(Roles);
		RoleList.setEditable(true);
		Office = new JTextField();	
		Dfa = new JTextField();
		LoginCred = new JTextField();
		ExtraField = new JTextField();
		ExtraField.setText("Insert Extra Field");
		ExtraField.setVisible(false);
		ExtraParam = new JTextField();
		ExtraParam.setText("Insert Extra Parameter");
		ExtraParam.setVisible(false);
		

//Radio Buttons and Corresponding Action Listeners are here
//------------------------------------------------------------
		UAEM = new JRadioButton("UA Format");
		USEM = new JRadioButton("US Format");
		UAEM.setSelected(true);
		
		UAEM.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            EmailUS = false;
	        }
	    });

	    //add disallow listener
	    USEM.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	EmailUS = true;
	        }
	    });
	    
	    EmailFormatGroup = new ButtonGroup();
	    EmailFormatGroup.add(UAEM);
	    EmailFormatGroup.add(USEM);
//-----------------------------------------------------
		ConfirmCompanyYes = new JRadioButton("Yes");
		ConfirmCompanyNo = new JRadioButton("No");
		ConfirmCompanyNo.setSelected(true);									
	
		ConfirmCompanyYes.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            Company.setEditable(true);
	        }
	    });

	    //add disallow listener
	    ConfirmCompanyNo.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            Company.setEditable(false);
	        }
	    });
	 
		CompanyGroup = new ButtonGroup();
		CompanyGroup.add(ConfirmCompanyYes);
		CompanyGroup.add(ConfirmCompanyNo);
//----------------------------------------------
		PhoneExtensionY = new JRadioButton("Required");
		PhoneExtensionN = new JRadioButton("Not Required");
		PhoneExtensionN.setSelected(true);

		PhoneExtensionY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            PhoneExtension.setEditable(true);
	            PhoneExtensionB = true;
	        }
	    });
	  	PhoneExtensionN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            PhoneExtension.setEditable(false);
	            PhoneExtensionB = false;
	            PhoneExtension.setText("");
	        }
	    });

	    PhoneExtensionGroup = new ButtonGroup();
	    PhoneExtensionGroup.add(PhoneExtensionY);
	    PhoneExtensionGroup.add(PhoneExtensionN);	
//-------------------------------------
		EmailGY = new JRadioButton("Required"); 
	    EmailGN = new JRadioButton("Not Required");
	    EmailGN.setSelected(true);

	    EmailGY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            EmailGroup.setEditable(true);
	            EmailGB = true;
	        }
	    });
	  	EmailGN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            EmailGroup.setEditable(false);
	            EmailGB = false;
	            EmailGroup.setText("");;
	        }
	    });

	    EmailGGroup = new ButtonGroup();;
	    EmailGGroup.add(EmailGY);
	    EmailGGroup.add(EmailGN);
//--------------------------------------------
		DomainADY = new JRadioButton("Required"); 
		DomainADN =  new JRadioButton("Not Required"); 
		DomainADN.setSelected(true);

		DomainADY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            DomainGroup.setEditable(true);
	            DomainADB = true;
	        }
	    });
	  	DomainADN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	  DomainGroup.setEditable(false);
	         	  DomainADB = false;
	         	  DomainGroup.setText("");
	        }
	    });

	    DomainGroupGroup = new ButtonGroup();
	    DomainGroupGroup.add(DomainADY);
	    DomainGroupGroup.add(DomainADN);
   //-----------------------------------------------------------
		EvoEAY  =  new JRadioButton("Required"); 
	 	EvoEAN  =  new JRadioButton("Not Required"); 
	 	EvoEAN.setSelected(true);

	 	EvoEAY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            EvoExtra.setEditable(true);	           
	        }
	    });
	  	EvoEAN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	  EvoExtra.setEditable(false);
	         	  EvoExtra.setText("");
	        }
	    });
	    EvoEAGroup = new ButtonGroup();
	    EvoEAGroup.add(EvoEAY);
	    EvoEAGroup.add(EvoEAN);
	   //-------------------------------------------------- 
	  	ClientEY =  new JRadioButton("Required"); 
	   	ClientEN =  new JRadioButton("Not Required"); 
	   	ClientEN.setSelected(true);

	   	ClientEY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {   
	        	ClientExtraList.setEditable(true);
	        }
	    });
	  	ClientEN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	  ClientExtraList.setEditable(false);
	         	  ClientExtraList.setSelectedIndex(0);
	          }
	    });
	    ClientEGroup = new ButtonGroup();
	    ClientEGroup.add(ClientEY);
	    ClientEGroup.add(ClientEN);
	    //------------------------------------------------- 
	    JiraY =  new JRadioButton("Required"); 
	    JiraN =  new JRadioButton("Not Required");
	    JiraN.setSelected(true); 
	    
	    JiraY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            Jira.setEditable(true);
	            JiraB = true;
	        }
	    });
	  	JiraN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	  Jira.setEditable(false);
	         	  JiraB = false;
	         	  Jira.setText("");
	        }
	    });
		JiraGroup = new ButtonGroup();
		JiraGroup.add(JiraY);
		JiraGroup.add(JiraN);	  		
//----------------------------------------------------
	    ConfY  =  new JRadioButton("Required"); 
	    ConfN =  new JRadioButton("Not Required"); 
	    ConfN.setSelected(true);

	     	ConfY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            Conf.setEditable(true);
	            ConfB = true;
	        }
	    });
	  	ConfN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	  Conf.setEditable(false);
	         	  ConfB = false;
	         	  Conf.setText("");
	        }
	    });
	  	ConfGroup = new ButtonGroup();
	  	ConfGroup.add(ConfY);
	  	ConfGroup.add(ConfN);
//-------------------------------------------------------
	    SvnY =  new JRadioButton("Required"); 
	    SvnN =  new JRadioButton("Not Required");
	    SvnN.setSelected(true); 
	    
	    SvnY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            Svn.setEditable(true);
	            SvnB = true;
	        }
	    });
	  	SvnN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	  Svn.setEditable(false);
	         	  SvnB = false;
	         	  Svn.setText("");
	        }
	    });
		SvnGroup = new ButtonGroup();
		SvnGroup.add(SvnY);
		SvnGroup.add(SvnN);	  		
//--------------------------------------------------
		BitBY =  new JRadioButton("Required"); 
		BitBN =  new JRadioButton("Not Required");
		BitBN.setSelected(true);

		 	BitBY.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	            BitBucket.setEditable(true);
	            BitB = true;
	        }
	    });
	  		BitBN.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	         	 BitBucket.setEditable(false);
	         	 BitB = false;
	         	 BitBucket.setText("");
	        }
	    });

	  	BitGroup = new ButtonGroup();
	  	BitGroup.add(BitBY);
	  	BitGroup.add(BitBN);
	  	
	  	//----------------------------------------------
	  	KeepFile =  new JRadioButton("Keep File"); 
		DeleteFile =  new JRadioButton("Delete File");
		DeleteFile.setSelected(true);
		
	  	 KeepFile.addActionListener(new ActionListener() {
		        @Override
		        public void actionPerformed(ActionEvent e) {
		          KeepEmailB = true;
		        }
		    });
		  DeleteFile.addActionListener(new ActionListener() {
		        @Override
		        public void actionPerformed(ActionEvent e) {
		         	KeepEmailB = false;
		        }
		    });

		SaveFileGroup = new ButtonGroup();
		SaveFileGroup.add(KeepFile);
		SaveFileGroup.add(DeleteFile);	  		

		
		//Submission of values into individual lines, and said lines into overline, here

		ConfirmSubmission = new JLabel();
		ConfirmSubmission.setVisible(false);
		ConfirmEmail = new JLabel();
		ConfirmEmail.setVisible(false);		
	  	
		Submit = new JButton("Submit");
		Submit.setVerticalTextPosition(AbstractButton.CENTER);
		
		AddExtra = new JButton("Submit Extra Parameter");
		AddExtra.setVisible(false);
		
		OpenExtra = new JButton("Add Extra Parameters");		
			
		Line2P1.add(TextName);
		Line2P2.add(FirstName);
		Line2P2.add(LastName);
		Line2P2.add(MiddleName);
		Line2.add(Line2P1);
		Line2.add(Line2P2);
				
		Line3P1.add(EmailT);
		Line3P2.add(UAEM);
		Line3P2.add(USEM);
		Line3P2.add(new JLabel(" "));
		Line3.add(Line3P1);
		Line3.add(Line3P2);		
		
		Line4.add(PhoneTextYNT);
		Line4.add(PhoneTextYNT);
		Line4.add(NumberList);
	/*
	Line4.add(PhoneExtensionY);
	Line4.add(PhoneExtensionN);
	Line4.add(PhoneExtension);
*/
		Line5P1.add(EmailGYNT);
		Line5P2.add(EmailGY);
		Line5P2.add(EmailGN);
		Line5P2.add(EmailGroup);
		Line5.add(Line5P1);
		Line5.add(Line5P2);
	
		Line7P1.add(DomainGroupYNT);
		Line7P2.add(DomainADY);
		Line7P2.add(DomainADN);
		Line7P2.add(DomainGroup);
		Line7.add(Line7P1);
		Line7.add(Line7P2);
	
		Line8P1.add(EvoExtraYNT);
		Line8P2.add(EvoEAY);
		Line8P2.add(EvoEAN);
		Line8P2.add(EvoExtra);
		Line8.add(Line8P1);
		Line8.add(Line8P2);
	
		Line9.add(AccessLevelT);
		Line9.add(AccessList);
	
		
		Line10P1.add(ClientExtraYNT);
		Line10P2.add(ClientEY);
		Line10P2.add(ClientEN);
		Line10P2.add(ClientExtraList);
		Line10.add(Line10P1);
		Line10.add(Line10P2);
	
		Line11.add(KayakoT);
		Line11.add(Kayako);

		Line12P1.add(JiraGT);
		Line12P2.add(JiraY);
		Line12P2.add(JiraN);
		Line12P2.add(Jira);
		Line12.add(Line12P1);
		Line12.add(Line12P2);
	
		Line13P1.add(ConfGT);
		Line13P2.add(ConfY);
		Line13P2.add(ConfN);
		Line13P2.add(Conf);
		Line13.add(Line13P1);
		Line13.add(Line13P2);
	
		Line14P1.add(SvnYNT);
		Line14P2.add(SvnY);
		Line14P2.add(SvnN);
		Line14P2.add(Svn);
		Line14.add(Line14P1);
		Line14.add(Line14P2);

		Line15P1.add(BitBucketYNT);
		Line15P2.add(BitBY);
		Line15P2.add(BitBN);
		Line15P2.add(BitBucket);
		Line15.add(Line15P1);
		Line15.add(Line15P2);
	
		Line16.add(BambooT);
		Line16.add(Bamboo);
	
		Line21.add(BambooRole);
		Line21.add(RoleList);
	
		Line17.add(OfficeT);
		Line17.add(Office);
	
		Line18.add(DfaT);
		Line18.add(Dfa);	

		Line19.add(LoginCredT);
		Line19.add(LoginCred);

		Line20.add(ConfirmSubmission);
		Line20.add(ConfirmEmail);
	
		Line22P1.add(KeepFileT);
		Line22P2.add(KeepFile);
		Line22P2.add(DeleteFile);
		Line22P2.add(new JLabel(" "));
		Line22.add(Line22P1);
		Line22.add(Line22P2);
	
		Line23.add(ExtraField);
		Line23.add(ExtraParam);
		Line23.add(AddExtra);
		Line23.add(OpenExtra);
	
		OverLine.add(Line2);
		OverLine.add(Line3);
		OverLine.add(Line4);
		OverLine.add(Line5);
		///OverLine.add(Line6);
		OverLine.add(Line7);
		OverLine.add(Line8);
		OverLine.add(Line9);
		OverLine.add(Line10);
		OverLine.add(Line11);
		OverLine.add(Line12);
		OverLine.add(Line13);
		OverLine.add(Line14);
		OverLine.add(Line15);
		OverLine.add(Line16);
		OverLine.add(Line21);
		OverLine.add(Line17);
		OverLine.add(Line18);
		OverLine.add(Line19);
		OverLine.add(Line22);
		OverLine.add(Line23);
		OverLine.add(Submit);
		OverLine.add(Line20);				
		OverLine.setPreferredSize(new Dimension(880, 590));
		Scroller = new JScrollPane(OverLine);
		OverLine.setAutoscrolls(true);
		Scroller.setPreferredSize(new Dimension(900,567));
		this.add(Scroller);

		//Backend which writes stuff to Excel here
		
		AddExtra.addActionListener(new ActionListener() {
		 public void actionPerformed(ActionEvent event) {
			 if(extraparams.equals("  ")) {
				 extraparams = "";
				 extrafields = "";
				 extraparams = extraparams + ExtraParam.getText() + " ";
				 extrafields = extrafields + "," + ExtraField.getText() + " ";
				
			 }
			 else { 
				 extraparams = extraparams + ", " + ExtraParam.getText() + " ";
				 extrafields = extrafields + ", " + ExtraField.getText() + " ";			
			 	  }				
				 ExtraParam.setText("  ");
				 ExtraField.setText("  ");
			}
		});
		OpenExtra.addActionListener(new ActionListener() {
		 public void actionPerformed(ActionEvent event) {	
			if(extravisible == false) {
				ExtraParam.setVisible(true);
				ExtraField.setVisible(true);
				AddExtra.setVisible(true);
				OpenExtra.setText("Revoke Extra Parameters");
				extravisible = true;
			}
			else {
				extraparams = "  ";
				extrafields = "  ";
				ExtraParam.setVisible(false);
				ExtraField.setVisible(false);
				AddExtra.setVisible(false);
				OpenExtra.setText("Add Extra Parameters");
				extravisible = false;
			  }
			}
		});
		
		Submit.addActionListener(new ActionListener() {
			 public void actionPerformed(ActionEvent event) {
				//Key
				//* = Mandatory
				//% = Optional
				 
				 ConfirmSubmission.setText("...Generating Excel File...");			
		         ConfirmSubmission.setVisible(true);   
				
				 FirstNameS = FirstName.getText() + "."; //* 
				 LastNameS = LastName.getText() + ".";
				 MiddleNameS = MiddleName.getText() + ".";
				 
				 PhoneExtensionS = (String) NumberList.getSelectedItem() + " ";
				 EmailGroupS = EmailGroup.getText() + ".";  //%
				 DomainGroupS = DomainGroup.getText() + "."; //%
				 EvoExtraS = EvoExtra.getText() + "."; //%
				 AccessLevelS = (String) AccessList.getSelectedItem() + " ";
				 ClientExtraS = (String) ClientExtraList.getSelectedItem() + " ";
				 KayakoS = Kayako.getText() + "."; //*
				 JiraS = Jira.getText() + "."; //%
				 ConfS = Conf.getText() + "."; //%
				 SvnS = Svn.getText() + "."; //%
				 BitBucketS = BitBucket.getText() + "."; //%
				 BambooS = Bamboo.getText() + "."; //*
				 OfficeS = Office.getText() + "."; //*
				 DfaS = Dfa.getText() + "."; //*
				 LoginCredS = LoginCred.getText() + "."; //*
				 String BS = (String) RoleList.getSelectedItem() + " ";
				 if(EmailUS == true) {
					 EmailS = FirstNameS.substring(0, 1) + LastNameS + "@evogence.com " ;
				 }
				 else {
					 EmailS = FirstNameS.substring(0, 1) + MiddleNameS.substring(0, 1) + LastNameS.substring(0, 1) + "@evogence.com ";
				 }
				 
				 
				 if (FirstNameS.equals(".") || LastNameS.equals(".") || EmailS.equals(".") || AccessLevelS.equals(".") 
						 || KayakoS.equals(".") || BambooS.equals(".") || OfficeS.equals(".") || DfaS.equals(".") || 
						 LoginCredS.equals(".") || (PhoneExtensionB == true && PhoneExtensionS.equals(".")) || 
						 (EmailGB == true && EmailGroupS.equals(".")) ||  (DomainADB == true && DomainGroupS.equals(".")) || 
						 (JiraB == true && JiraS.equals(".")) || (ConfB == true && ConfS.equals(".")) || 
						 (SvnB == true && SvnS.equals(".")) ||  (BitB == true && BitBucketS.equals(".")))
						 {
		    				info = new  ArrayList<String>(Arrays.asList(FirstNameS, LastNameS, MiddleNameS, 
								EmailS, EvoExtraS, AccessLevelS, KayakoS, BambooS, 
								BS, OfficeS, DfaS, LoginCredS));
						
						String Require = ("First Name, Last Name, Middle, Email, Evo Extranet, Access Level, Kayako, Bamboo, Bamboo Level, Office, DfaS, Login Credentials");
						require = new ArrayList<String>(Arrays.asList(Require.split(",")));
					 	String hold = "";
					 	for(int i = 0; i < info.size(); i++) {
					 	if(info.get(i).equals(".")) {
					 			hold = hold + require.get(i) + " "; 
					 		}
					 	}
					 //	ConfirmSubmission.setVerticalTextPosition(500);
					 	ConfirmSubmission.setText("Error - Fill Out the following Fields: " + hold);
					 	ConfirmSubmission.setVisible(true);
					 	}
					 	
				 
			 else { 
				  	
				 	String Require = ("First Name,Last Name,Middile Initial,Email,PhoneExtension,Email Group,Domain Group,Evo Extranet,Access Level,Client Extranet,"
			 			+ "Kayako,Jira,CONF,SVN,Bitbucket,Bamboo,Bamboo Level,Microsoft OFfice,DFA,Login Credentials" + extrafields);
			    	SendEmail = true;			 	
			    	require = new ArrayList<String>(Arrays.asList(Require.split(",")));

			    	extraparamstring = new ArrayList<String>(Arrays.asList(extraparams.split(",")));
			    	info = new  ArrayList<String>(Arrays.asList(FirstNameS, LastNameS, MiddleNameS, EmailS, PhoneExtensionS, EmailGroupS, DomainGroupS, EvoExtraS, AccessLevelS, ClientExtraS, KayakoS, JiraS, ConfS, SvnS, BitBucketS, BambooS, 
						BS, OfficeS, DfaS, LoginCredS));
			    	for(int i = 0; i < extraparamstring.size(); i++) {
					info.add(extraparamstring.get(i));
			    	}
			    	
				
			
			    	String filename = "Insert file directory here" + LastName.getText() + ".xls" ;			
			    	HSSFWorkbook workbook = new HSSFWorkbook();
			    	HSSFSheet sheet = workbook.createSheet("User Info Sheet");  
			    	
			    	HSSFRow rowhead = sheet.createRow((short)0);
			    		for(int i = 0; i < require.size(); i++) {
			    			rowhead.createCell(i).setCellValue(require.get(i));
			    		}
			    	HSSFRow row = sheet.createRow((short)1);
			    		for(int i = 0; i < info.size(); i++) {
			    			String filler = info.get(i).substring(0, info.get(i).length()-1);
			    			if(info.get(i).length() == 1) {filler="N/A";}
			    			row.createCell(i).setCellValue(filler);     
			    		}
				
			    	FileOutputStream fileOut;
			    	try {
					 fileOut = new FileOutputStream(filename);
					 workbook.write(fileOut);
			         fileOut.close();
			    	} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
			    	} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
			    	}
			    	try {
					workbook.close();
			    	} 
			    	catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
			    	}
	           
	            ConfirmSubmission.setText("Your excel file has been generated!");			
	            ConfirmSubmission.setVisible(true);   
	          }
				 if(SendEmail == true) {
					 				
				    final String username = "insert email address to send email here";
					final String password = "insert password for said email address";

					Properties props = new Properties();
					props.put("mail.smtp.auth", "true");
					props.put("mail.smtp.starttls.enable", "true");
					props.put("mail.smtp.host", "smtp.gmail.com");
					props.put("mail.smtp.port", "587");

					Session session = Session.getInstance(props,
					  new javax.mail.Authenticator() {
						protected PasswordAuthentication getPasswordAuthentication() {
							return new PasswordAuthentication(username, password);
						}
					  });

					try {

						Message message = new MimeMessage(session);
						message.setFrom(new InternetAddress("insert same email address seen in username variable.com"));
						message.setRecipients(Message.RecipientType.TO,
							InternetAddress.parse("insert email address you will send this to"));
						message.setSubject("New File Upload");
						BodyPart messageBodyPart = new MimeBodyPart();
						messageBodyPart.setText("Hello, you have recieved a new file");
						Multipart multipart = new MimeMultipart();				         
				        multipart.addBodyPart(messageBodyPart);
				        
				         messageBodyPart = new MimeBodyPart();
				         filename = "insert path here" + LastName.getText() + ".xls";		
				         System.out.println(filename);
				        
				         DataSource source = new FileDataSource(filename);
				         messageBodyPart.setDataHandler(new DataHandler(source));
				         messageBodyPart.setFileName(LastName.getText() + " Profile.xls");
				      
				         multipart.addBodyPart(messageBodyPart);
				         message.setContent(multipart);
 						Transport.send(message);
    					ConfirmEmail.setText("Email Sent!");
	   					ConfirmEmail.setVisible(true);
	   				
					} catch (MessagingException e) {
						ConfirmEmail.setText("Error sending email");
						ConfirmEmail.setVisible(true);
						throw new RuntimeException(e);
					}
				if(KeepEmailB == false){
					try {
						filename = "insert path here" + LastName.getText() + ".xls";	//This line is For testing so I don't get spammed with emails, delete once finished.
					    path = Paths.get(filename);
						Files.delete(path);
					} catch (NoSuchFileException x) {
					    System.err.format("%s: no such" + " file or directory%n", path);
					} catch (DirectoryNotEmptyException x) {
					    System.err.format("%s not empty%n", path);
					} catch (IOException x) {
					    // File permission problems are caught here.
					    System.err.println(x);
		   		   }
			     } 
			   }
			 } 
		 });
			
	}
	
}
