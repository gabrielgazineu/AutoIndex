import java.io.File;
import java.io.IOException;
import java.util.Locale;

import jxl.CellFeatures;
import jxl.CellView;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class autoIndex {

	private String inputFile;
	private String outputFile;
	private String ListaEntSai;
	private static WritableCellFormat arialNoBoldNoUnderline;
	private static WritableCellFormat arialBoldNoUnderline;


	public void setOutputFile(String inputFile, String outputFile, String entSai) {
		this.inputFile = inputFile;
		this.outputFile = outputFile;
		this.ListaEntSai = entSai;
	}

	public void createExcel() throws IOException, WriteException, BiffException {
		File inFile = new File(this.inputFile);
		File outFile = new File(this.outputFile);
		File entSai = new File(this.ListaEntSai);
		WorkbookSettings wbSettings = new WorkbookSettings();
		wbSettings.setLocale(new Locale("en", "EN"));
		Workbook workbookInput = Workbook.getWorkbook(inFile);
		Sheet inputSheet = workbookInput.getSheet(0);
		WritableWorkbook workbookOutput = Workbook.createWorkbook(outFile, wbSettings);
		workbookOutput.createSheet("Instrument Index", 0);
		WritableSheet outputSheet = workbookOutput.getSheet(0);
		Workbook workbookEntSai = Workbook.getWorkbook(entSai);
		Sheet entSaiSheet = workbookEntSai.getSheet(0);
		
		autoIndex.prepareHeaders(inputSheet, outputSheet);
		
		//TODO Chamar função para ler documento de entradas e saidas

		workbookOutput.write();
		workbookOutput.close();
	}

	private static void addCaption(WritableSheet sheet, int column, int row, String s, boolean bold)
			throws RowsExceededException, WriteException {

		WritableFont arial8ptNoBoldNoUnderline = new WritableFont(WritableFont.ARIAL, 8, WritableFont.NO_BOLD, false,UnderlineStyle.NO_UNDERLINE);
		arialNoBoldNoUnderline = new WritableCellFormat(arial8ptNoBoldNoUnderline);
		arialNoBoldNoUnderline.setAlignment(Alignment.CENTRE);
		arialNoBoldNoUnderline.setVerticalAlignment(VerticalAlignment.TOP);
		

		WritableFont arial8ptBoldNoUnderline = new WritableFont(WritableFont.ARIAL, 8, WritableFont.BOLD, false,UnderlineStyle.NO_UNDERLINE);
		arialBoldNoUnderline = new WritableCellFormat(arial8ptBoldNoUnderline);
		arialBoldNoUnderline.setAlignment(Alignment.CENTRE);
		arialBoldNoUnderline.setVerticalAlignment(VerticalAlignment.TOP);

		if (bold) {
			Label label = new Label(column, row, s, arialBoldNoUnderline);
			sheet.addCell(label);
		} else {
			Label label = new Label(column, row, s, arialNoBoldNoUnderline);
			sheet.addCell(label);
		}

	}


	public static void prepareHeaders(Sheet input, WritableSheet output){

		String cellContent;
		CellView columnView;
		
		CellView rowView = input.getRowView(0);

		try {
			output.setRowView(0, rowView);
			for (int i = 0; i < 77; i++) {
				//Set cell content
				cellContent = input.getCell(i, 0).getContents();
				addCaption(output, i, 0, cellContent, false);
								
				
				//Set Column width
				columnView = input.getColumnView(i);
				output.setColumnView(i, columnView);
				
				
			}
		} catch (WriteException e) {			
			e.printStackTrace();
		}
	}


	public static void main(String[] args) throws WriteException, IOException, BiffException {

		String directoryOutput = "C:/Users/Gabriel M.G/AutoIndex/workspace/AutoIndex/Teste.xls";
		String directoryInput = "C:/Users/Gabriel M.G/AutoIndex/workspace/AutoIndex/U_51-Instrument_Index.xls";
		String directoryEntSai = "C:/Users/Gabriel M.G/AutoIndex/workspace/AutoIndex/U_51-Lista_Entrada_Saida.xls";

		autoIndex index = new autoIndex();
		index.setOutputFile(directoryInput, directoryOutput, directoryEntSai);
		index.createExcel();
		System.out.println("Please check the result file under: " + directoryOutput);
	}
}