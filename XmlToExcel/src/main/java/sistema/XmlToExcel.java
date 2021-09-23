package sistema;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import com.fincatto.documentofiscal.nfe400.classes.nota.NFNota;
import com.fincatto.documentofiscal.nfe400.classes.nota.NFNotaProcessada;
import com.fincatto.documentofiscal.utils.DFPersister;


public class XmlToExcel {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

	Boolean criaExcel = false;
	Boolean leExcel = false;
	Boolean digitaEan = false;
	Boolean criaXml = false;
	String caminho;
	File xml = null;
	File[] xmls = null;
	
	NFNota nota = null;
	
	if (args.length == 0) {
		
		System.out.println("Favor informar os parametros necessarios em caso de duvida digitar HELP");
		System.exit(0);
	} else if (args[0].contains("help")) {
		System.out.println("  -c                  Parametro usado para informar o caminho do XML -c C:\\users\\1020.xml");
		System.out.println("  -l                  Parametro para importar arquivo excell com as modificaçoes");
		System.out.println("  -d                  Parametro usado o uso do Robo para digitação do EAN no sistema este parametro tem que ser usado junto com o parametro -c ou -l");
		System.out.println("  -x                  Parametro usado para substituir os dados da nota pelo do excell este parametro tem que ser usado junto com o parametro -l");
		
		System.exit(0);
	} else {
		
		for (int i = 0; i < args.length; i++) {
			
			switch (args[i]) {
			case "-c":
					try {
						i++;
						if (args[i].contains(".xml") || args[i].contains(".XML") ) {	
							xml = new File(args[i]);
							criaExcel = true;
						} else {
							File diretorio = new File(args[i]);
							diretorio.isDirectory();
							xmls = diretorio.listFiles(new FileFilter() {
								
								@Override
								public boolean accept(File pathname) {
									return pathname.getName().endsWith(".xml");
								}
							});
							criaExcel = true;
						}
					} catch (Exception e) {
						System.out.println(e.toString());
						System.exit(0);
					}		
			break;
			case "-l":
				leExcel = true;
			break;
			case "-d":
				digitaEan = true;
			break;
			case "-x":
				criaXml = true;
			break;
			}
		}
	}
	

	List<Produto> produtos = new ArrayList<Produto>();
	
	System.out.println(System.getProperty("user.home").toString()+"\\Desktop\\novo.xslx");
	caminho = System.getProperty("user.home").toString()+"\\Desktop\\novo.xslx";
	System.out.println(caminho);
	
	
	if (criaExcel) {
		
		if (xml==null) {
			for (int i = 0; i < xmls.length; i++) {
				NFNotaProcessada notaProcessada = new DFPersister(false).read(NFNotaProcessada.class,xmls[i]);
				nota = notaProcessada.getNota();
				String nfNumero = nota.getInfo().getIdentificacao().getNumeroNota().toString();
				nota.getInfo().getItens().forEach(iten->{ 
					Integer quantidade;
					if(iten.getProduto().getQuantidadeComercial().contains(".")) {
						quantidade = Integer.valueOf(iten.getProduto().getQuantidadeComercial().substring(0, iten.getProduto().getQuantidadeComercial().indexOf(".")));
					} else {
						quantidade =  Integer.valueOf(iten.getProduto().getQuantidadeComercial());
					}
					produtos.add(new Produto(nfNumero,iten.getProduto().getCodigo(),iten.getProduto().getDescricao(),
							iten.getProduto().getNcm(), new BigDecimal(iten.getProduto().getValorUnitario()).setScale(2, RoundingMode.HALF_EVEN),
							quantidade,iten.getProduto().getCodigoDeBarras()
							));
				});
			}
		}else {
			
			NFNotaProcessada notaProcessada = new DFPersister(false).read(NFNotaProcessada.class,xml);
			nota = notaProcessada.getNota();
			String nfNumero = nota.getInfo().getIdentificacao().getNumeroNota().toString();
			nota.getInfo().getItens().forEach(iten->{ 
				Integer quantidade;
				if(iten.getProduto().getQuantidadeComercial().contains(".")) {
					quantidade = Integer.valueOf(iten.getProduto().getQuantidadeComercial().substring(0, iten.getProduto().getQuantidadeComercial().indexOf(".")));
				} else {
					quantidade =  Integer.valueOf(iten.getProduto().getQuantidadeComercial());
				}
				produtos.add(new Produto(nfNumero,iten.getProduto().getCodigo(),iten.getProduto().getDescricao(),
						iten.getProduto().getNcm(), new BigDecimal(iten.getProduto().getValorUnitario()).setScale(2, RoundingMode.HALF_EVEN),
						quantidade,iten.getProduto().getCodigoDeBarras()
						));
			});
		}

		criaExcell(produtos, caminho);
	}
	
	System.out.println("Digite enter para continuar");
	System.in.read();
	
	if (leExcel) {
		
		FileInputStream excel = new FileInputStream(caminho);
		HSSFWorkbook pastaDeTrabalho = new HSSFWorkbook(excel);
		HSSFSheet planilha = pastaDeTrabalho.getSheet("Produtos");
		produtos.clear();
		planilha.forEach(linha ->{
			Produto produto = new Produto();
			linha.forEach(celula ->{
				switch (celula.getColumnIndex()) {
				case 1:
					produto.setCod(pegaValorCelula(celula));
				break;
				case 2:
					produto.setDescricao(pegaValorCelula(celula));
				break;
				case 5:
					produto.setQuant(Integer.getInteger(pegaValorCelula(celula)));
				break;
				case 6:
					produto.setEan(pegaValorCelula(celula));
				break;
				}
				
			});
			System.out.println(produto.toString());
			produtos.add(produto);			
			
		});
		if(produtos.get(0).getCod().equals("Codigo")) {
			produtos.remove(0);
		}
		if (criaXml) {
			nota.getInfo().getItens().forEach(iten ->{
				
				iten.getProduto().setCodigo(produtos.get(0).getCod());
				iten.getProduto().setDescricao(produtos.get(0).getDescricao());
				if(!(produtos.get(0) == null) ||!produtos.get(0).getEan().isEmpty()) {
					iten.getProduto().setCodigoDeBarras(produtos.get(0).getEan());
				}
				produtos.remove(0);
			});
			String texto = nota.toString();
			BufferedWriter bw = new BufferedWriter(new FileWriter(new File("C:\\Users\\fabio\\Desktop\\nota1.xml")));
			bw.write(texto);
			bw.close();
		}
	}
	
	if(digitaEan) {
		if(produtos.get(0).getCod().equals("Codigo")) {
			produtos.remove(0);
		}
		TimeUnit.SECONDS.sleep(10); 
		Comandos comando = new Comandos(); 
		comando.digitaComando(produtos);
		}
		
	}
	private static void criaExcell(List<Produto> produtos, String caminho) {
		try {
			HSSFWorkbook pastaDeTrabalho = new HSSFWorkbook(); 
			HSSFSheet novaPlanilha = pastaDeTrabalho.createSheet("Produtos");
			
			
			int numeroLinha = 0;
			String[] cabecalho = {"Nota","Codigo","Descrição","NCM","Valor de Compra","Quantidade","EAN"};
			Row linha = novaPlanilha.createRow(numeroLinha++);
			for (int i = 0; i < cabecalho.length; i++) {
				Cell dados = linha.createCell(i);
				dados.setCellValue(cabecalho[i]);
			}			
			for (Produto produto : produtos) {
				linha = novaPlanilha.createRow(numeroLinha++);
				int numeroCelula = 0;
				Cell cellNota = linha.createCell(numeroCelula++);
				cellNota.setCellValue(produto.getNfNumero());
				Cell cellCod = linha.createCell(numeroCelula++);
				cellCod.setCellValue(produto.getCod());
				Cell cellDesc = linha.createCell(numeroCelula++);
				cellDesc.setCellValue(produto.getDescricao());
				Cell cellNcm = linha.createCell(numeroCelula++);
				cellNcm.setCellValue(produto.getNcm());
				Cell cellValor = linha.createCell(numeroCelula++);
				cellValor.setCellValue(produto.getValorDeCompra().doubleValue());
				Cell cellQuant = linha.createCell(numeroCelula++);
				cellQuant.setCellValue(produto.getQuant());
				if  (produto.getEan()== null || produto.getEan().isEmpty()) {
					Cell cellEan = linha.createCell(numeroCelula++);
					cellEan.setCellValue(produto.getEan());					
				}
				
			}
			FileOutputStream out = new FileOutputStream(caminho);
			pastaDeTrabalho.write(out);
			out.close();
			pastaDeTrabalho.close();
			System.out.println("Arquivo Criado");
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo não encontrado");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("erro na edição do arquivo");
		}
	}
	
	private static String  pegaValorCelula (Cell celula) {
		
		if (celula == null) {
			return "";
		}
		CellType tipo =  celula.getCellType();
		switch(tipo) {
		case NUMERIC:
            return String.valueOf(celula.getNumericCellValue());
        case STRING:
            return celula.getStringCellValue();
        case BOOLEAN:
            return String.valueOf(celula.getBooleanCellValue());
        case BLANK:
            return "";
        case ERROR:
            return String.valueOf(celula.getErrorCellValue());
        default:
            throw new RuntimeException("Tipo de celula nao esperado (" + tipo + ")");
		
		}
		
	}

	
	

}
