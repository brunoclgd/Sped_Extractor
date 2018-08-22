import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.OutputStream;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpedExtractor {
	private static int index = 0;
	private static int line1 = 1;
	private static int line2 = 1;
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		//Scanner ler = new Scanner(System.in);
		
		int fileSize = 0;
		try {

			File arquivos[] = null;
			File diretorio = new File("C:\\Users\\teste\\Desktop\\novos");
			arquivos = diretorio.listFiles();
			String extensionToFind = ".txt";
			int qtd_arquivos = arquivos.length;
			for(int k = 0; k < arquivos.length; k++) {
				System.out.println("***ARQUIVO "+k+1+" de "+qtd_arquivos+"***");
				System.out.println(arquivos[k]);
				fileSize = arquivos[k].getName().length();
				
				if(arquivos[k].getName().substring(fileSize-4, fileSize).equals(extensionToFind)) {
					index = 0;
					line1 = 1;
					line2 = 1;
					Workbook wb = new XSSFWorkbook();
					//OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/"+arquivos[k].getName().substring(0,  46)+".xlsx");
					OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/EFD CONTRIBUIÇÕES 06 2018 c170.xlsx");
					FileReader arq = new FileReader(arquivos[k]);
					
					String content[];

					Sheet sheet = wb.createSheet("0200");
					Sheet sheet2 = wb.createSheet("c170");
					Row rowTitle1 = sheet.createRow(0);
					Row rowTitle2 = sheet2.createRow(0);

					rowTitle1.createCell(0).setCellValue("CÓDIGO DO ITEM");
					rowTitle1.createCell(1).setCellValue("DESCRIÇÃO DO ITEM");
					rowTitle1.createCell(2).setCellValue("CÓDIGO DA NCM");

					rowTitle2.createCell(0).setCellValue("CÓDIGO DO ITEM");
					rowTitle2.createCell(1).setCellValue("QUANTIDADE DO ITEM");
					rowTitle2.createCell(2).setCellValue("UNIDADE DO ITEM");
					rowTitle2.createCell(3).setCellValue("VALOR DO ITEM");
					rowTitle2.createCell(4).setCellValue("VALOR DO DESCONTO COMERCIAL");
					rowTitle2.createCell(5).setCellValue("CST_ICMS");
					rowTitle2.createCell(6).setCellValue("CFOP");
					rowTitle2.createCell(7).setCellValue("VALOR DA BASE DE CALC. ICMS");
					rowTitle2.createCell(8).setCellValue("ALÍQUOTA DO ICMS");
					rowTitle2.createCell(9).setCellValue("VALOR DO ICMS CRED/DEB");
					rowTitle2.createCell(10).setCellValue("VALOR DA BCST");
					rowTitle2.createCell(11).setCellValue("ALÍQUOTA DO ICMSST");
					rowTitle2.createCell(12).setCellValue("VALOR DO ICMSST");
					rowTitle2.createCell(13).setCellValue("CÓDIGO ST-PIS");
					rowTitle2.createCell(14).setCellValue("VALOR BC-PIS");
					rowTitle2.createCell(15).setCellValue("ALÍQUOTA DO PIS(%)");
					rowTitle2.createCell(16).setCellValue("QTD. BC-PIS");
					rowTitle2.createCell(17).setCellValue("VALOR DO PIS");
					rowTitle2.createCell(18).setCellValue("CST-COFINS");
					rowTitle2.createCell(19).setCellValue("VALOR DA BC-COFINS");
					rowTitle2.createCell(20).setCellValue("ALÍQUOTA DO COFINS(%)");
					rowTitle2.createCell(21).setCellValue("QTD. BC-COFINS");
					rowTitle2.createCell(22).setCellValue("VALOR DA COFINS");

					BufferedReader lerArq = new BufferedReader(arq);
					String linha = lerArq.readLine();

					while(linha != null && linha.length() > 0) {
						if(!linha.isEmpty()) {
							linha = linha.replace("|", ";");
							linha = linha.substring(1, linha.length());
							content = linha.split(";");
							if(content != null && content[0].equalsIgnoreCase("0200")) {
								Row rowData = sheet.createRow(line1);
								System.out.println("CÓDIGO DO ITEM: "+content[1]);
								rowData.createCell(0).setCellValue(content[1]);
								sheet.autoSizeColumn(0);
								System.out.println("DESCRIÇÃO DO ITEM: "+content[2]);
								rowData.createCell(1).setCellValue(content[2]);
								sheet.autoSizeColumn(1);
								System.out.println("CÓDIGO DA NCM: "+content[7]);
								rowData.createCell(2).setCellValue(content[7]);
								sheet.autoSizeColumn(2);
								line1++;

							}
							else if(content != null && content[0].equalsIgnoreCase("c170")) {
								if(content[24].equals("04") || content[24].equals("06") || 
										content[30].equals("04") || content[30].equals("06")) {
									Row rowData2 = sheet2.createRow(line2);
									System.out.println("CÓDIGO DO ITEM: "+content[2]);
									rowData2.createCell(0).setCellValue(content[2]);
									sheet2.autoSizeColumn(0);

									System.out.println("QUANTIDADE DO ITEM: "+content[4]);
									rowData2.createCell(1).setCellValue(content[4]);
									sheet2.autoSizeColumn(1);

									System.out.println("UNIDADE DO ITEM: "+content[5]);
									rowData2.createCell(2).setCellValue(content[5]);
									sheet2.autoSizeColumn(2);

									System.out.println("VALOR DO ITEM: R$ "+content[6]);
									rowData2.createCell(3).setCellValue(content[6]);
									sheet2.autoSizeColumn(3);

									System.out.println("VALOR DO DESCONTO COMERCIAL: R$ "+content[7]);
									rowData2.createCell(4).setCellValue(content[7]);
									sheet2.autoSizeColumn(4);

									System.out.println("CST_ICMS: "+content[9]);
									rowData2.createCell(5).setCellValue(content[9]);
									sheet2.autoSizeColumn(5);

									System.out.println("CFOP: "+content[10]);
									rowData2.createCell(6).setCellValue(content[10]);
									sheet.autoSizeColumn(6);

									System.out.println("VALOR DA BASE DE CALC. ICMS: "+content[12]);
									rowData2.createCell(7).setCellValue(content[12]);
									sheet2.autoSizeColumn(7);

									System.out.println("ALÍQUOTA DO ICMS: "+content[13]);
									rowData2.createCell(8).setCellValue(content[13]);
									sheet2.autoSizeColumn(8);

									System.out.println("VALOR DO ICMS CRED/DEB: "+ content[14]);
									rowData2.createCell(9).setCellValue(content[14]);
									sheet2.autoSizeColumn(9);

									System.out.println("VALOR DA BCST: "+content[15]);
									rowData2.createCell(10).setCellValue(content[15]);
									sheet2.autoSizeColumn(10);

									System.out.println("ALÍQUOTA DO ICMSST: "+content[16]);
									rowData2.createCell(11).setCellValue(content[16]);
									sheet2.autoSizeColumn(11);

									System.out.println("VALOR DO ICMSST: "+content[17]);
									rowData2.createCell(12).setCellValue(content[17]);
									sheet2.autoSizeColumn(12);

									System.out.println("CÓDIGO ST-PIS: "+content[24]);
									rowData2.createCell(13).setCellValue(content[24]);
									sheet2.autoSizeColumn(13);

									System.out.println("VALOR BC-PIS "+content[25]);
									rowData2.createCell(14).setCellValue(content[25]);
									sheet2.autoSizeColumn(14);

									System.out.println("ALÍQUOTA DO PIS(%): "+content[26]);
									rowData2.createCell(15).setCellValue(content[26]);
									sheet2.autoSizeColumn(15);

									System.out.println("QTD. BC-PIS: "+content[27]);
									rowData2.createCell(16).setCellValue(content[27]);
									sheet2.autoSizeColumn(16);

									System.out.println("VALOR DO PIS: "+content[29]);
									rowData2.createCell(17).setCellValue(content[29]);
									sheet2.autoSizeColumn(17);

									System.out.println("CST-COFINS: "+content[30]);
									rowData2.createCell(18).setCellValue(content[30]);
									sheet2.autoSizeColumn(18);

									System.out.println("VALOR DA BC-COFINS: "+content[31]);
									rowData2.createCell(19).setCellValue(content[31]);
									sheet2.autoSizeColumn(19);

									System.out.println("ALÍQUOTA DO COFINS(%): "+content[32]);
									rowData2.createCell(20).setCellValue(content[32]);
									sheet2.autoSizeColumn(20);

									System.out.println("QTD. BC-COFINS: "+content[33]);
									rowData2.createCell(21).setCellValue(content[33]);
									sheet2.autoSizeColumn(21);

									System.out.println("VALOR DA COFINS: "+content[34]);
									rowData2.createCell(22).setCellValue(content[34]);
									sheet2.autoSizeColumn(22);
									line2++;
								}
								
							}
						}
						
						
						linha = lerArq.readLine();
						

					}
					lerArq.close();
					wb.write(fileOut);
					fileOut.close();
					wb.close();
				}
				
			}
		
		
		
	}catch(Exception e) {
		System.out.println(e.getLocalizedMessage());
		System.out.println(e.getCause());
		System.out.println(e);
	}


}

}
