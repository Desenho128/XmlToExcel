package sistema;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.TextArea;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import javax.swing.JTextField;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JTextArea;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.awt.event.ActionEvent;
import javax.swing.JScrollPane;

public class InterfaceRobo extends JFrame {

	private JPanel contentPane;
	private JTextField textField;
	private JTextField textField_1;

	/**
	 * Create the frame.
	 */
	public InterfaceRobo() {
		getContentPane().setLayout(null);
		
		textField_1 = new JTextField();
		textField_1.setBounds(23, 26, 86, 20);
		getContentPane().add(textField_1);
		textField_1.setColumns(10);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		textField = new JTextField();
		textField.setBounds(10, 11, 227, 20);
		contentPane.add(textField);
		textField.setColumns(10);
		
		JButton btnNewButton = new JButton("Planilha");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			
				JFileChooser selecionar = new JFileChooser();
				selecionar.setFileSelectionMode(JFileChooser.FILES_ONLY);
				selecionar.setApproveButtonText("Selecionar");
				selecionar.setDialogTitle("Selecionar Arquivo");
				selecionar.setFileFilter(new FileFilter() {
					
					@Override
					public boolean accept(File f) {
						if (f.getName().endsWith(".xslx")) {
							return true;
						} else {
						return false;
						}
					}

					@Override
					public String getDescription() {
						// TODO Auto-generated method stub
						return "XSLX";
					}
				});
				int i = selecionar.showSaveDialog(null);
				if (i==1) {
					textField.setText("");
				} else {
					textField.setText(selecionar.getSelectedFile().getPath());
					
				}

				
			}
		});
		btnNewButton.setBounds(247, 10, 89, 23);
		contentPane.add(btnNewButton);
		
		JScrollPane scrollPane = new JScrollPane();
		scrollPane.setBounds(10, 50, 414, 200);
		contentPane.add(scrollPane);
		
		JTextArea textArea = new JTextArea();
		scrollPane.setViewportView(textArea);
		textArea.setLineWrap(true);
		textArea.setWrapStyleWord(true);

		
		JButton btnNewButton_1 = new JButton("Digitar");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			List<Produto> produtos =	importarExcel(textField.getText(),textArea);
			if(produtos.get(0).getCod().equals("Codigo")) {
				produtos.remove(0);
			}
			try {
				produtos.forEach(produto ->{
					textArea.append(produto.toString()+"\n");
					textArea.updateUI();
				});
				
				for(int i =0; i < 11; i++) { 
					TimeUnit.SECONDS.sleep(1);
					textArea.append(String.valueOf(i+"\n"));
				}
				Comandos comando = new Comandos(); 
				comando.digitaComando(produtos);
			} catch (Exception e2) {
				
			}

			}
		});
		btnNewButton_1.setBounds(343, 10, 81, 23);
		contentPane.add(btnNewButton_1);
		
	}
	
	
	private List importarExcel (String caminho,JTextArea textarea) {
		
		try (FileInputStream excel = new FileInputStream(caminho)) {
			HSSFWorkbook pastaDeTrabalho = new HSSFWorkbook(excel);
			HSSFSheet planilha = pastaDeTrabalho.getSheet("Produtos");
			List<Produto> produtos = new ArrayList<Produto>();
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
						if (pegaValorCelula(celula).equals("Quantidade")) {
							produto.setQuant(0);
						} else if (pegaValorCelula(celula).contains(".")) {
							produto.setQuant(Integer.valueOf(pegaValorCelula(celula).substring(0,pegaValorCelula(celula).indexOf("."))));
							}else {
							produto.setQuant(Integer.valueOf(pegaValorCelula(celula)));
						}
					break;
					case 6:
						produto.setEan(pegaValorCelula(celula));
					break;
					}
					
				});
				System.out.println(produto.toString());
				produtos.add(produto);			
				
			});
			return produtos;

			
		} catch (Exception e) {
			System.out.println(e.toString());
			
			return null;
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
