package PrimaryPackage;

import java.awt.EventQueue;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JButton;

import java.awt.BorderLayout;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.Buffer;

import javax.swing.JTextArea;
import javax.swing.text.AbstractDocument.Content;

import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.POIXMLProperties.CustomProperties;
import org.apache.poi.POIXMLProperties.ExtendedProperties;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class SecondPack {
	
	int headingsCount = 0;
	

	private JFrame frame;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					SecondPack window = new SecondPack();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public SecondPack() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		final JTextArea textArea = new JTextArea();
		
		JButton btnOpenFile = new JButton("Open File");
		btnOpenFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.out.println("button clicked!");
				
				String finalString = "<html>";
				
				finalString =  finalString.concat("<head> <title> Sample Section Document </title> </head> <body>");
				
				
				try {
					
					
					JFileChooser chooser = new JFileChooser();
					chooser.showOpenDialog(null);
					final XWPFDocument doc = new XWPFDocument(new FileInputStream(chooser.getSelectedFile()));
					XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
					
					//CoreProperties s = extractor.getCoreProperties();
					//CustomProperties s = extractor.getCustomProperties();
					ExtendedProperties s = extractor.getExtendedProperties();
					
					System.out.println(s.toString());
					//System.out.println(s.getUnderlyingProperties());
					
			       // textArea.setText(extractor.getText());
					
					int size = doc.getParagraphs().size();
					
					doc.getBodyElements().size();
					
					
                    int i=0;
                    String headingTitleForComparison = "Heading1";
					
					for(i=0;i<size;i++)
					{
						String fetchingStringForHeadingCount = doc.getParagraphArray(i).getStyle();
						if(headingTitleForComparison.equals(fetchingStringForHeadingCount))
						{
							headingsCount++;
						}
					}
					
					final int headingCountPositions[] = new int[headingsCount];
					int endOfPagePositions[] = new int[headingsCount];
					
					int j=0;
					for(i=0;i<size;i++)
					{
						String fetchingStringForHeadingCount = doc.getParagraphArray(i).getStyle();
						if(headingTitleForComparison.equals(fetchingStringForHeadingCount))
						{
							headingCountPositions[j] =  i;
							j++;
							System.out.println(i + " ");
						}
					}
					
//					j=0;
//					int k=0;
//					int l=0;
//					int m=0;
//					for(i=headingCountPositions[k];(k < headingsCount) && i<headingCountPositions.length;i=headingCountPositions[k++])
//					{
//						l = headingCountPositions[k];
//						//String fetchingStringForHeadingCount = doc.getParagraphArray(i).getStyle();
//						while( (k < headingsCount) && l<headingCountPositions[k+1])
//						{
//							if((null == doc.getParagraphArray(i).getStyle()) && !"".equalsIgnoreCase(doc.getParagraphArray(i).getText()))
//							{
//								endOfPagePositions[m] = l;
//								System.out.println(endOfPagePositions[m]);
//								m++;
//								break;
//							}
//							
//							l++;
//						}
//						
//					}
					
					
					
					new Thread() {
			            public void run() {
			                /* block of code which need to execute via thread */
			            	
			            	int headingCountPointer = 0;
							int currentPosition = 0;
							String tentativeTitle = new String();// = new Strin();
							
							int status = 0;
							
							while (status < headingsCount)
							{
								String finalTempString = "<html> <head> <title> Sample Section Document </title> </head> <body>" +"\n";
								
								System.out.println("Current pos = "+headingCountPositions[headingCountPointer]);
								
								//(headingsCount > headingCountPointer+1) &&
								
								
								
								 while(currentPosition < headingCountPositions[headingCountPointer])
								{
									 System.out.println("while loop: " + currentPosition);
									if("Heading1".equalsIgnoreCase(doc.getParagraphArray(currentPosition).getStyle()))
									{
							//			finalString = finalString.concat("<p> <b> "+ doc.getParagraphArray(0).getText() + "</b> </p>" ); 
										finalTempString = finalTempString.concat("<p> <b> "+ doc.getParagraphArray(currentPosition).getText() + "</b> </p>" +"\n");
									}						
									else if("Subtitle".equalsIgnoreCase(doc.getParagraphArray(currentPosition).getStyle()))
									{
										//tentativeTitle = new String("Section"+doc.getParagraphArray(currentPosition).getText()+".html");
										tentativeTitle = "Section"+doc.getParagraphArray(currentPosition).getText()+".html";
									}
									else if((null == doc.getParagraphArray(currentPosition).getStyle()) && !"".equalsIgnoreCase(doc.getParagraphArray(currentPosition).getText()))
									{
										//finalString = finalString.concat("<p> "+ doc.getParagraphArray(currentPosition).getText() + " </p>" ); 
										finalTempString = finalTempString.concat("<p> "+ doc.getParagraphArray(currentPosition).getText() + " </p>" +"\n" ); 
									}
									else if("ListParagraph".equalsIgnoreCase(doc.getParagraphArray(currentPosition).getStyle()))
									{
										//finalString = finalString.concat("<ul> <li> "+ doc.getParagraphArray(currentPosition).getText() + " </li> </ul>" );
										finalTempString = finalTempString.concat("<ul> <li> "+ doc.getParagraphArray(currentPosition).getText() + " </li> </ul>" +"\n" );
									}
									currentPosition++;
								}
									
									
								finalTempString = finalTempString.concat("</body> </html>");
								
								if(currentPosition == 0)
								{
									
								}
								else
								{
									File htmlFileToCreateAndWrite = new File("C:/temp/JavaTempWorkspace/WindowsBuilderProTemp/"+"file"+Integer.toString(currentPosition)+".html");
									if (!htmlFileToCreateAndWrite.exists()) {
										try {
											htmlFileToCreateAndWrite.createNewFile();
										} catch (IOException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
									}
									
									FileWriter fileWriterForHTMLFile;
									try {
										fileWriterForHTMLFile = new FileWriter(htmlFileToCreateAndWrite.getAbsoluteFile());
										BufferedWriter bufferedWriterForHTMLFile = new BufferedWriter(fileWriterForHTMLFile);
										bufferedWriterForHTMLFile.write(finalTempString);
										bufferedWriterForHTMLFile.close();
									} catch (IOException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									
									
									
									System.out.println("File Written "+Integer.toString(headingCountPointer));
								}
								
								
								finalTempString = "";
								
								
								headingCountPointer++;
								status++;
							}
			            	
			            }
			        }.start();
					
					
					
					
					//finalString = finalString.concat("<p> <b> "+ doc.getParagraphArray(0).getText() + "</b> </p>" );
					
					
					textArea.setText(doc.getParagraphArray(6).getText() + "\n \n \n \n------------------\n \n \n \n" + 
			        doc.getParagraphArray(6).getStyle() +  "\n \n------------------\n \n" + Integer.toString(size) 
			        + "\n \n------------------\n \n" + 
			        doc.getProperties().getExtendedProperties().getUnderlyingProperties().getPages()
							+ "\n \n------------------\n \n" + Integer.toString(headingsCount));
					 
					
					 
					 //textArea.setText();
					 
					//FileOutputStream theHTMLFile = new FileOutputStream("section_1a_b.html");
					//theHTMLFile.write();
					
					
					
					
					
					
					finalString = finalString.concat("</body> </html>");
					
					File htmlFileToCreateAndWrite = new File("section_1a_b.html");
					FileWriter fileWriterForHTMLFile = new FileWriter(htmlFileToCreateAndWrite.getAbsoluteFile());
					BufferedWriter bufferedWriterForHTMLFile = new BufferedWriter(fileWriterForHTMLFile);
		            bufferedWriterForHTMLFile.write(finalString); 
					bufferedWriterForHTMLFile.close();
					
					System.out.println("File Written");
					
				} catch (FileNotFoundException e) {
					
					
					// TODO Auto-generated catch block
					e.printStackTrace();
					
					
				} catch (IOException e) {
					
					
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
			}
		});
		frame.getContentPane().add(btnOpenFile, BorderLayout.NORTH);
		
		
		frame.getContentPane().add(textArea, BorderLayout.CENTER);
	}

}
