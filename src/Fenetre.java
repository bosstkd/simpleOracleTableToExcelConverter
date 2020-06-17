import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;


public class Fenetre extends JFrame {

	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			
			
			
			
			
			public void run() {
				
				 try {
		        	 UIManager.setLookAndFeel("com.jtattoo.plaf.mcwin.McWinLookAndFeel");        	
		        } catch (ClassNotFoundException e) {
				} catch (InstantiationException e) {
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				} catch (UnsupportedLookAndFeelException e) {
					e.printStackTrace();
				}
				
				try {
					Fenetre frame = new Fenetre();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	
	JTextField tf_url = new JTextField();
	private JTextField tf_nom_u;
	private JPasswordField pf_pwd;

	
	
	public Fenetre() {
		setResizable(false);
		setTitle("Excel to Oracle Converter v 1.0");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 424, 276);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JLabel lblAdresseIp = new JLabel("URL");
		lblAdresseIp.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		lblAdresseIp.setBounds(48, 11, 83, 29);
		contentPane.add(lblAdresseIp);
		
		JLabel lblNomDutilisateur = new JLabel("Nom D'utilisateur");
		lblNomDutilisateur.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		lblNomDutilisateur.setBounds(47, 78, 135, 22);
		contentPane.add(lblNomDutilisateur);
		
		JLabel lblMotDePasse = new JLabel("Mot de passe");
		lblMotDePasse.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		lblMotDePasse.setBounds(48, 128, 123, 29);
		contentPane.add(lblMotDePasse);
		tf_url.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		
		
		tf_url.setBounds(289, 10, 107, 29);
		contentPane.add(tf_url);
		
		tf_nom_u = new JTextField();
		tf_nom_u.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		tf_nom_u.setBounds(289, 70, 107, 38);
		contentPane.add(tf_nom_u);
		tf_nom_u.setColumns(10);
		
		pf_pwd = new JPasswordField();
		pf_pwd.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		pf_pwd.setBounds(289, 123, 107, 38);
		contentPane.add(pf_pwd);
		
		JButton btnConfirmer = new JButton("Confirmer");
		btnConfirmer.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				if(tf_nom_u.getText().equals("")||tf_nom_u.getText().equals(" ")||tf_url.getText().equals("")||tf_url.getText().equals(" ")||pf_pwd.getText().equals("")||pf_pwd.getText().equals("")){
                     JOptionPane.showMessageDialog(null,"Verifier les champs SVP !!","Attention",JOptionPane.WARNING_MESSAGE);
				}else{
					 try {
		                 // Chargement du pilote JDBC
		                 Class.forName("oracle.jdbc.driver.OracleDriver");
		                 // URL de connexion
		                 String url = "jdbc:oracle:thin:@//"+tf_url.getText()+":1521/XE";
		                 String user = tf_nom_u.getText();
		                 String password = pf_pwd.getText();
		                 // Connexion
		                 Connection con = null;
		                     try {
		                         con = DriverManager.getConnection(url, user, password);
			                     JOptionPane.showMessageDialog(null,"Connextion établie","Information",JOptionPane.INFORMATION_MESSAGE);

			                     
			                     essaye2();
			                     
			                     
		                     } catch (SQLException ex) {
		                         Logger.getLogger(Fenetre.class.getName()).log(Level.SEVERE, null, ex);
		 	                     JOptionPane.showMessageDialog(null,"Erreur : " + ex,"Erreur",JOptionPane.ERROR_MESSAGE);

		                     }


		    		 
		                       con.close();
		             
		                      } catch (ClassNotFoundException e1) {
		                    	  
		 	                     JOptionPane.showMessageDialog(null,"Erreur lors du chargement du pilote : " + e1,"Erreur",JOptionPane.ERROR_MESSAGE);

		                     } catch (SQLException sqle) {
		                    	 
		 	                     JOptionPane.showMessageDialog(null,"Erreur SQL : " + sqle,"Erreur",JOptionPane.ERROR_MESSAGE);

		                     } 
				}
				
				
				

			}
		});
		
		
		btnConfirmer.setFont(new Font("Segoe Print", Font.PLAIN, 13));
		btnConfirmer.setBounds(232, 172, 117, 46);
		contentPane.add(btnConfirmer);
		
		JLabel lblNewLabel = new JLabel("creer par Mahmoudi Med ElAmine");
		lblNewLabel.setFont(new Font("Tahoma", Font.PLAIN, 9));
		lblNewLabel.setBounds(10, 225, 148, 23);
		contentPane.add(lblNewLabel);
		
		JLabel lblNewLabel_1 = new JLabel("Contact us : a-ek@hotmail.fr");
		lblNewLabel_1.setFont(new Font("Tahoma", Font.PLAIN, 9));
		lblNewLabel_1.setBounds(285, 227, 123, 19);
		contentPane.add(lblNewLabel_1);
	}
	
//************************************************
	
	
	void essaye2() {
		final JFileChooser chooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
		    "Fichier d'extension .xlsx", "xlsx");
		chooser.setFileFilter(filter);	
		
		
		int returnVal = chooser.showOpenDialog(chooser);
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			
			
			
		try {
			FileChannel fileSource;
			fileSource = new FileInputStream(chooser.getSelectedFile().getAbsolutePath()).getChannel();

			
			//System.out.println(chooser.getSelectedFile().getAbsolutePath());

			String table = chooser.getSelectedFile().getName().substring(0,chooser.getSelectedFile().getName().length()-5);
			
			c(chooser.getSelectedFile().getAbsolutePath() , table);
			
			
			
			
			   JOptionPane.showMessageDialog(null, "Conversion terminer .");
			  
			   
			   if (fileSource != null) {
			       try {
					fileSource.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			   }
			   
			   
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
										   }
	   
		} 
	}
	
	
	
	void c(String chemin , String table){
		
		Vector vct = new Vector();
		Vector vct2 = new Vector();
		Vector vct3 = new Vector();
		
		table = table.toUpperCase().replaceAll(" ", "_");
		
		String req = "CREATE TABLE "+table+"(";
		
		
		String reqDeb = "INSERT INTO "+table+"( ";
		
		String reqSec = "";

		String reqFin = "";
		
		String req1 = "";

		
		 XSSFWorkbook wb;
		try {
			wb = new XSSFWorkbook(chemin);
			   XSSFSheet sheet = (XSSFSheet)wb.getSheetAt(0);
			   XSSFRow row = sheet.getRow(0);
			   XSSFCell cell = row.getCell((short) 0);
	       int i = 0;          
           boolean ts = true;
			for (Iterator rowIt = sheet.rowIterator(); rowIt.hasNext();) {   
	            for (Iterator cellIt = row.cellIterator(); cellIt.hasNext();) {
	            
	              cell = (XSSFCell) cellIt.next();	     
	              
	              
	              
	              if(ts)i++;
	              if(cell.getCellType()==1){
	            	  String x = cell.getStringCellValue().toUpperCase().replaceAll("'", ""); 
		              if(cell.getRowIndex()==0 && ts){
		            	  vct.addElement(x);
		              }
		              else if(cell.getRowIndex()==1){
		            	  vct2.addElement(x);
		              }
	              }
	              if(cell.getRowIndex()>1){
	            	  if(cell.getCellType()==0){
	            		  float x = (float) cell.getNumericCellValue();
	            		  vct3.addElement(x);
	            	  }else
	            		  if(cell.getCellType()==3){
   	    	            	  System.out.println("valeurs de cellule : "+cell.getStringCellValue());
	            			  vct3.addElement(cell.getStringCellValue());
	                     }
	            		  else
	                		 if (cell.getCellType()==1){
	    	            	  String x = cell.getStringCellValue().toUpperCase().replaceAll("'", ""); 
	    	            	  System.out.println("valeurs de x : "+x);
	            			  vct3.addElement(x);
	            		      } 
	              } 	              
	          }     	            
	            if(ts) ts = false;
	            
	            row = (XSSFRow) rowIt.next();
	            
	            if(cell.getRowIndex()==2){
	            	
	            	boolean bl = true;
	            	
	            	
	            	for(int j = 0; j < vct.size(); j++){
	            		//System.out.println(vct.elementAt(j)+"\t");
	            		try{
	            			vct.elementAt(j+1);
	            		}catch(Exception e){
	            			bl = false;
	            		}
	            		if(bl){
	            			req = req + vct.elementAt(j)+" "+vct2.elementAt(j)+", ";
	            			reqSec = reqSec + vct.elementAt(j)+", ";
	            		}else{
	            			req = req + vct.elementAt(j)+" "+vct2.elementAt(j)+")";
	            			reqSec = reqSec + vct.elementAt(j)+") VALUES (";
	            		}
	            	}
	            	
	            	
	            	requette(req, table);
	            	
	            	
	            	
	            	
	            	
	            System.out.println(req);	
	            }
	            
	            if(cell.getRowIndex()>=2){
	            	
	            	boolean bl = true;
	            	
	            	//boolean num = false;
	            	float test_f = 0;

	            	for(int j = 0; j < vct.size(); j++){
	            		
	            		try{
	            			vct.elementAt(j+1);
	            		}catch(Exception e){
	            			bl = false;
	            		}
	            		
	            		
	            		
	            		
	            		try{
	            			test_f = Float.parseFloat((String) vct3.elementAt(j));
	            		
	            			if(bl){
		            			reqFin = reqFin + vct3.elementAt(j)+" , ";
		            		}else{
		            			reqFin = reqFin + vct3.elementAt(j)+" )";
		            		}
	            			
	            		}catch(Exception e){
	            			
	            			if(bl){
		            			reqFin = reqFin +"'"+ vct3.elementAt(j)+"' , ";
		            		}else{
		            			reqFin = reqFin +"'"+ vct3.elementAt(j)+"' )";
		            		}
	            		}
	            		
	            		
	            		
	            		
	            	}
	            	
	            	
	            	req1 = reqDeb + reqSec + reqFin;
		            
		            System.out.println(req1);
		            
		            requette(req1);
		            
		            reqFin = "";
		            
		            req1 = "";
		            
	      		vct3.removeAllElements();	
      		  }
	            
	                        
	            
	          }
			
			
	//*******************************		
			

			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
		
	}
	
	
//************************************************	
	
	
void requette(String req, String table){
	try {
        // Chargement du pilote JDBC
        Class.forName("oracle.jdbc.driver.OracleDriver");
        // URL de connexion
        String url = "jdbc:oracle:thin:@//"+tf_url.getText()+":1521/XE";
        String user = tf_nom_u.getText();
        String password = pf_pwd.getText();
        // Connexion
        Connection con = null;
            try {
                con = DriverManager.getConnection(url, user, password);
            } catch (SQLException ex) {
                Logger.getLogger(Fenetre.class.getName()).log(Level.SEVERE, null, ex);
            }

            
        	Statement statement = con.createStatement();
        	
        	
        	String del = "drop TABLE "+table+" ";
			try{
        	statement.executeUpdate(del);
			statement.executeUpdate(req);

			}catch(Exception e){
				statement.executeUpdate(req);
			}
              con.close();
             } catch (ClassNotFoundException e1) {
           	  
                 JOptionPane.showMessageDialog(null,"Erreur lors du chargement du pilote : " + e1,"Erreur",JOptionPane.ERROR_MESSAGE);

            } catch (SQLException sqle) {
           	 
                 JOptionPane.showMessageDialog(null,"Erreur SQL : " + sqle,"Erreur",JOptionPane.ERROR_MESSAGE);

            } 
}
	

void requette(String req){
	try {
        Class.forName("oracle.jdbc.driver.OracleDriver");
        String url = "jdbc:oracle:thin:@//"+tf_url.getText()+":1521/XE";
        String user = tf_nom_u.getText();
        String password = pf_pwd.getText();
        Connection con = null;
            try {
                con = DriverManager.getConnection(url, user, password);
            } catch (SQLException ex) {
                Logger.getLogger(Fenetre.class.getName()).log(Level.SEVERE, null, ex);
            }
        	Statement statement = con.createStatement();

        	statement.executeUpdate(req);
        	
        	try {
				Thread.sleep(800);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        	
			
              con.close();
             } catch (ClassNotFoundException e1) {
           	  
                 JOptionPane.showMessageDialog(null,"Erreur lors du chargement du pilote : " + e1,"Erreur",JOptionPane.ERROR_MESSAGE);

            } catch (SQLException sqle) {
           	 
                 JOptionPane.showMessageDialog(null,"Erreur SQL : " + sqle,"Erreur",JOptionPane.ERROR_MESSAGE);

            } 
}
}
