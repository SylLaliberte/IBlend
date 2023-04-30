/****************************************************************
	 * Ingredient blender										*
	 * by Sylvain Laliberté										*										
	 * April 28, 2023											*
	 * v1.0														*
	 *															*
	 *À ajouter													*
	 *- flag pour regrouper certains ingrédients				*
	 *															*
	 *															*
	 *															*
	 ************************************************************/

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IBlend implements ActionListener, ItemListener {
	
	// Variable declaration
	
	JFrame frame;
	JTextArea textfield;
	JButton[] functionButtons = new JButton[3];
	JButton OpenButton, CopyButton, SaveButton;
	JPanel panel;
	JComboBox<String> comboBox;
	FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Spreadsheet",".xslx");
	JFileChooser fileChooser;
	File file;
	
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	int qtyOfSheet;
	
	XSSFWorkbook outWorkbook;
	XSSFSheet outSheet;
	
	String [] sheets;
	ArrayList<String> sheetsInFile = new ArrayList<String>(qtyOfSheet);
	ImageIcon image = new ImageIcon(".\\data\\icon3.png");
	int rows;
	int cols;
	
	Object [][] data = new Object[rows][cols]; 	// creation of 2d Object array
	Object [][] newdata = new Object[rows][cols]; // new 2d array for results	
	
	ArrayList <Double> subtotal = new ArrayList<Double>();
	ArrayList <String> ingredient = new ArrayList<String>();
	HashMap<String, Double> map = new HashMap<>(); 
	HashMap<String, Double> sortedMap = new HashMap<>();
	
	List<Double> sortedValues;
	List<String> sortedIng = new ArrayList<>();
	
	Font myFont = new Font("Verdana", Font.PLAIN, 15);
	
	File fileSaved;
	
	// frame constructor
	IBlend() {
		// Blender frame
		frame = new JFrame("Ingredient Blender");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(640,480);
		frame.setLayout(null);
		
		//Blender window
		textfield = new JTextArea();
		textfield.setBounds(25,60,575,300);
		textfield.setFont(myFont);
		textfield.setEditable(false);
		textfield.setBackground(Color.WHITE);
		textfield.setLineWrap(true);
		textfield.setWrapStyleWord(true);
		
		//Blender buttons
		OpenButton = new JButton("Open");
		CopyButton = new JButton("Copy");
		SaveButton = new JButton("Save");
		
		
		functionButtons[0]= OpenButton;
		functionButtons[1]= CopyButton;
		functionButtons[2]= SaveButton;
		
		for(int i=0;i<3;i++) {
			functionButtons[i].addActionListener(this);
			functionButtons[i].setFont(myFont);
			functionButtons[i].setFocusable(false);
		}
		
		OpenButton.setBounds(25,375,150,50);
		CopyButton.setBounds(245,375,150,50);
		SaveButton.setBounds(450,375,150,50);
		
		//Drop down list for sheets in file
	
		comboBox = new JComboBox<String>();
		comboBox.addItemListener(this);
		comboBox.setBounds(25, 15, 575, 25);
		comboBox.setFont(myFont);
		
		//frame.add(panel);
		frame.add(OpenButton);
		frame.add(CopyButton);
		frame.add(SaveButton);
		frame.add(textfield);
		frame.add(comboBox);
		center(frame);
		frame.setVisible(true);
		
		frame.setIconImage(image.getImage());
		
	}  //End IBlend
	
	public static void main(String[] args) {
		
		IBlend blender = new IBlend();

	} // End Main

	//Action performed by buttons
	@Override
	public void actionPerformed(ActionEvent e) {
	
		if(e.getSource()==OpenButton) {
			try {
				file = openAFile(fileChooser);
			} catch (InvalidFormatException | IOException e1) {
				System.out.println("Warning:Exception in format");
				e1.printStackTrace();
			}
			System.out.println(file);
		
		}
		
		if(e.getSource()==CopyButton) {
			StringSelection stringSelection = new StringSelection (textfield.getText());
			Clipboard clpbrd = Toolkit.getDefaultToolkit ().getSystemClipboard ();
			clpbrd.setContents (stringSelection, null);
		}
	
		if(e.getSource()==SaveButton) {
			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setCurrentDirectory(new File("."));  // add address to folder
			
			int responseSave = fileChooser.showSaveDialog(null); //select file to open
			System.out.println("response: "+responseSave);
			
			if(responseSave == JFileChooser.APPROVE_OPTION)
			{
				fileSaved = new File(fileChooser.getSelectedFile().getAbsolutePath());
				System.out.println(fileSaved);
				outWorkbook = populatingOutWorkBookwithList(sortedIng, sortedValues, rows, cols);
				FileOutputStream outstream = null;
				try {
					outstream = new FileOutputStream(fileSaved);
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} // open the output stream
				try {
					outWorkbook.write(outstream);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} // write file
				try {
					outWorkbook.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				try {
					outstream.close();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}  	
			}
		}
		
	}  // End Action performed
	
	// Item selected in ComboBox
	public void itemStateChanged (ItemEvent e)
	{
	if (e.getSource () == comboBox) {
		
		sortedIng.clear();
		textfield.selectAll();
	    textfield.replaceSelection("");
		System.out.println(comboBox.getSelectedItem());
		System.out.println(comboBox.getSelectedIndex());
		
		sheet = workbook.getSheetAt(comboBox.getSelectedIndex());
		
		rows=sheet.getLastRowNum()+1;				//returns # of rows  ??Problem, 1 missing...
		cols=sheet.getRow(0).getLastCellNum();		//returns number of columns at row 0
		System.out.println("size array, rows "+rows+" x "+"cols "+cols);
		data = workbookToDataArray(workbook,sheet, rows, cols);						
		try {
			workbook.close();
		} catch (IOException e1) {
			System.out.println("Warning: IOException");
			e1.printStackTrace();
		}
		
		// calculation of ing% of each item in mix
		newdata = getIngValueInMix(data,rows, cols); 
		
		// Creation of ingredient list and subtotal of each
		ingredient = databaseIngExtraction(newdata, rows, cols);					
		subtotal = databaseSubTExtraction(newdata, rows, cols);	
					
		//Create a map.  Match ingredient with its subtotal
		map = createAMapFromTwoList(ingredient,subtotal);	
		sortedValues = mapSortingUtil(map);
		textfield.setText(sortedIng.toString());
		}
		
	}

/****************** FUNCTIONS *************************** FUNCTIONS ************************* FUNCTIONS ***********************/
	
	// Frame centering function
	public static void center(JFrame frame) {
		 
        // get the size of the screen
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
 
        // calculate the new location of the window
        int w = frame.getSize().width;
        int h = frame.getSize().height;
 
        int x = (dim.width - w) / 2;
        int y = (dim.height - h) / 2;
 
        // moves this component to a new location, the top-left corner of
        // the new location is specified by the x and y
        // parameters in the coordinate space of this component's parent
        frame.setLocation(x, y);
 
    }
	
	// Open a file with JFilechooser
	private File openAFile(JFileChooser fileChooser) throws InvalidFormatException, IOException {
 		fileChooser = new JFileChooser();
 		fileChooser.addChoosableFileFilter(filter);
		fileChooser.setCurrentDirectory(new File("."));  // add address to folder
		int response = fileChooser.showOpenDialog(null); //select file to open
		System.out.println("response: "+response);
		
		if(response == JFileChooser.APPROVE_OPTION)
		{
			file = new File(fileChooser.getSelectedFile().getAbsolutePath());
			//System.out.println(file);
									
			workbook = new XSSFWorkbook(file);
			int qtyOfSheets = workbook.getNumberOfSheets();
			sheets = new String[qtyOfSheets];
			for(int i=0;i<qtyOfSheets;i++) {
				sheets[i]=workbook.getSheetName(i);
				comboBox.addItem(sheets[i]);
				System.out.println(sheets[i]);
			}
			System.out.println(qtyOfSheet);
		}
		return file;
	}
	
	// excel to DataArray
	private static Object[][] workbookToDataArray(XSSFWorkbook workbook, XSSFSheet sheet, int rows, int cols) {
		
		Object [][] data = new Object[rows][cols]; 	// creation of 2d Object array
		
		for(int r=0;r<rows;r++)					
		{
			XSSFRow row=sheet.getRow(r);			// point row number
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell=row.getCell(c);		// point cell number
				
				//protection against empty cell.  Set to 0.
				if (cell != null) 
				{	
					switch(cell.getCellType())
					{
						case STRING: data[r][c]=cell.getStringCellValue();break;
						case NUMERIC: data[r][c]=cell.getNumericCellValue();break;
						case BOOLEAN: data[r][c]=cell.getBooleanCellValue();break;
						default:
							break;
					}
				}
				else
				{
					data[r][c]=0;
				}
				//System.out.print(data[r][c]+"  ");
			}
			//System.out.println();
		}
		return data;
	}
	
	// Calculation of ingredient proportion in mix
	public static Object[][] getIngValueInMix(Object[][] data, int rows, int cols) {
	
		Object [][] newdata = new Object[rows][cols];
		for(int r=0;r<rows;r++)					
		{
			for(int c=0;c<cols;c++)
			{
				if(data[r][c] instanceof String) 
					newdata[r][c]=data[r][c];  // if cell is a String, copy data in new cell
				if(data[r][c] instanceof Double) 
					newdata[r][c]= (double)data[r][c]*(double)data[r][cols-1]; //if cell's a number, multiply cell by last column
				//System.out.print(newdata[r][c]+"  "); //print new array
			}
			//System.out.println();
		}
		return newdata;
		} 
	
	// Ingredient name extraction
	private static ArrayList<String> databaseIngExtraction(Object[][] newdata, int rows, int cols) {
		ArrayList <String> ingredient = new ArrayList<String>();
		String ing= new String();
		
		for(int c=1;c<cols;c++)
		{
			for(int r=0;r<rows;r++)
			{
				ing = (String)newdata[0][c];
			}
			ingredient.add(ing);
		}
		return ingredient;
	}
	
	// Ing. value subtotal calculation and extraction
	private static ArrayList<Double> databaseSubTExtraction(Object[][] newdata, int rows, int cols) {
		ArrayList <Double> subtotal = new ArrayList<Double>();
		
		for(int c=1;c<cols;c++)
		{
			double sum=0;
			
			for(int r=0;r<rows;r++)
			{	
				if(newdata[r][c] instanceof Double) 
					sum=sum+(double)newdata[r][c];	
			}
			subtotal.add(sum);
		}
		return subtotal;
	}
	
	// Creation of map
	private static HashMap<String, Double> createAMapFromTwoList(ArrayList<String> ingredient, ArrayList<Double> subtotal) {
		HashMap<String, Double> map = new HashMap<>();
		
		for(int i = 0; i < ingredient.size()-1; i++) 
		{
		// sort out the ingredient at 0%
			if(subtotal.get(i) != 0) {
		    	map.put(ingredient.get(i), subtotal.get(i));
		    }
			
		}
			map.entrySet().stream()
				.sorted(Map.Entry.comparingByValue(Comparator.reverseOrder()))
				.forEach(System.out::println);
			
			return map;
	}

	// Map sorting and creation of lists
	private List<Double> mapSortingUtil(HashMap<String, Double> map2) {
		
		List<Double> sortedValues = map2.entrySet().stream()
	            //sort a Map by value and stored in resultSortedValue
			.sorted(Map.Entry.comparingByValue(Comparator.reverseOrder()))
	        .peek(e -> sortedIng.add(e.getKey()))
	        .map(x -> 100*x.getValue())	//Multiplying by 100 to get more readable %	           
	        .collect(Collectors.toList());
		return sortedValues;
	}

	// Excel file from 2 sorted lists
	
	// Creation of final workbook
	public static XSSFWorkbook populatingOutWorkBookwithList(List<String> sortedIng, List<Double> sortedValues,int rows, int cols) 
	{
		XSSFWorkbook outWorkbook2 = new XSSFWorkbook();  //create new workbook
		XSSFSheet outSheet2=outWorkbook2.createSheet("sorted"); //create and name sheet
		
		for(int r=0;r<2;r++) {
		
		XSSFRow row=outSheet2.createRow(r); // create row (r)
		
		if (row.getRowNum()==0) 
		{
			for(int c=0;c<sortedIng.size();c++) {
				XSSFCell cell=row.createCell(c); //create cell (c)
				cell.setCellValue(sortedIng.get(c));
				}		
		}	
		if (row.getRowNum()==1)
		{
			for(int c=0;c<sortedIng.size();c++) {
				XSSFCell cell=row.createCell(c); //create cell (d)
				cell.setCellValue(sortedValues.get(c));
				}	
		}
		}
		return outWorkbook2;
	}
	
} //End Class
	


