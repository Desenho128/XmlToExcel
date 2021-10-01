package sistema;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JInternalFrame;
import javax.swing.JButton;
import javax.swing.JFileChooser;

import java.awt.BorderLayout;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.fincatto.documentofiscal.nfe400.classes.nota.NFNota;
import com.fincatto.documentofiscal.nfe400.classes.nota.NFNotaProcessada;
import com.fincatto.documentofiscal.utils.DFPersister;

import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;
import java.awt.event.ActionEvent;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import java.awt.Font;

public class InterfaceXmlToExcel {

	public JFrame frame;
	private JTextField textField;
	private JButton btnNewButton_1;
	private JTextField textField_1;
	private JButton btnNewButton_4;


	/**
	 * Create the application.
	 */
	public InterfaceXmlToExcel() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.getContentPane().setFont(new Font("Tahoma", Font.BOLD, 11));
		frame.setBounds(100, 100, 556, 367);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JButton btnNewButton = new JButton("Arquivo");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				JFileChooser selecionar = new JFileChooser();
				selecionar.setFileSelectionMode(JFileChooser.FILES_ONLY);
				selecionar.setApproveButtonText("Selecionar");
				selecionar.setDialogTitle("Selecionar Arquivo");
				selecionar.setFileFilter(new FileFilter() {
					
					@Override
					public boolean accept(File f) {
						if (f.getName().endsWith(".xml")|| f.getName().endsWith(".XML")) {
							return true;
						} else {
						return false;
						}
					}

					@Override
					public String getDescription() {
						// TODO Auto-generated method stub
						return "XMLs";
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
		btnNewButton.setBounds(317, 83, 99, 23);
		frame.getContentPane().add(btnNewButton);
		
		textField = new JTextField();
		textField.setBounds(10, 84, 303, 20);
		frame.getContentPane().add(textField);
		textField.setColumns(10);
		
		btnNewButton_1 = new JButton("Diretorio");
		btnNewButton_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser selecionar = new JFileChooser();
				selecionar.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				selecionar.setApproveButtonText("Selecionar");
				selecionar.setDialogTitle("Selecionar Diretorio");
				if (1 == selecionar.showOpenDialog(null)) {
					textField.setText("");
				}else {
					textField.setText(selecionar.getSelectedFile().getPath());
					
				}
				
			}
		});
		btnNewButton_1.setBounds(426, 83, 99, 23);
		frame.getContentPane().add(btnNewButton_1);
		
		JLabel lblNewLabel = new JLabel("Caminho");
		lblNewLabel.setBounds(10, 65, 89, 14);
		frame.getContentPane().add(lblNewLabel);
		
		textField_1 = new JTextField();
		textField_1.setBounds(10, 118, 303, 20);
		frame.getContentPane().add(textField_1);
		textField_1.setColumns(10);
		
		JButton btnNewButton_2 = new JButton("Local");
		btnNewButton_2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser selecionar = new JFileChooser();
				selecionar.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				selecionar.setApproveButtonText("Selecionar");
				selecionar.setDialogTitle("Selecionar Diretorio");
				if (1 == selecionar.showOpenDialog(null)) {
					textField_1.setText("");
				}else {
					textField_1.setText(selecionar.getSelectedFile().getPath());
					
				}				
			}
		});
		btnNewButton_2.setBounds(317, 117, 99, 23);
		frame.getContentPane().add(btnNewButton_2);
		
		JButton btnNewButton_3 = new JButton("Converter ");
		btnNewButton_3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				if (textField.getText().trim().isEmpty())
					
					JOptionPane.showMessageDialog(null,"Selecionar Arquivo ou Pasta ");
				else if (textField_1.getText().trim().isEmpty()) {
					
					JOptionPane.showMessageDialog(null,"Selecionar local onde sera gravado o Arquivo ");
					
				} else {
					try {
						List<Produto> produtos = new ArrayList<Produto>();
						if (textField.getText().endsWith(".xml") || textField.getText().endsWith(".XML")) {
							File xml = new File(textField.getText());
							NFNotaProcessada notaProcessada = new DFPersister(false).read(NFNotaProcessada.class, xml);
							NFNota nota = notaProcessada.getNota();
							String nfNumero = nota.getInfo().getIdentificacao().getNumeroNota().toString();
							nota.getInfo().getItens().forEach(iten -> {
								Integer quantidade;
								if (iten.getProduto().getQuantidadeComercial().contains(".")) {
									quantidade = Integer.valueOf(iten.getProduto().getQuantidadeComercial().substring(0,
											iten.getProduto().getQuantidadeComercial().indexOf(".")));
								} else {
									quantidade = Integer.valueOf(iten.getProduto().getQuantidadeComercial());
								}
								produtos.add(new Produto(nfNumero, iten.getProduto().getCodigo(),
										iten.getProduto().getDescricao(), iten.getProduto().getNcm(),
										new BigDecimal(iten.getProduto().getValorUnitario()).setScale(2,
												RoundingMode.HALF_EVEN),
										quantidade, iten.getProduto().getCodigoDeBarras()));
							});
						} else {
							File[] xmls = null;
							File caminho = new File(textField.getText());
							xmls = caminho.listFiles(new java.io.FileFilter() {

								@Override
								public boolean accept(File f) {
									if (f.getName().endsWith(".xml") || f.getName().endsWith(".XML")) {
										return true;
									} else {
										return false;
									}
								}
							});
							for (int i = 0; i < xmls.length; i++) {
								NFNotaProcessada notaProcessada = new DFPersister(false).read(NFNotaProcessada.class,
										xmls[i]);
								NFNota nota = notaProcessada.getNota();
								String nfNumero = nota.getInfo().getIdentificacao().getNumeroNota().toString();
								nota.getInfo().getItens().forEach(iten -> {
									Integer quantidade;
									if (iten.getProduto().getQuantidadeComercial().contains(".")) {
										quantidade = Integer.valueOf(iten.getProduto().getQuantidadeComercial()
												.substring(0, iten.getProduto().getQuantidadeComercial().indexOf(".")));
									} else {
										quantidade = Integer.valueOf(iten.getProduto().getQuantidadeComercial());
									}
									produtos.add(new Produto(nfNumero, iten.getProduto().getCodigo(),
											iten.getProduto().getDescricao(), iten.getProduto().getNcm(),
											new BigDecimal(iten.getProduto().getValorUnitario()).setScale(2,
													RoundingMode.HALF_EVEN),
											quantidade, iten.getProduto().getCodigoDeBarras()));
								});
							}

						}

						criaExcell(produtos, textField_1.getText() + "\\novo.xslx");

					} catch (Exception e2) {
						// TODO: handle exception
					}
				}
			}
		});
		btnNewButton_3.setBounds(176, 216, 167, 46);
		frame.getContentPane().add(btnNewButton_3);
		
		btnNewButton_4 = new JButton("Importar Excel");
		btnNewButton_4.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			
				InterfaceRobo robo = new InterfaceRobo();
				robo.setVisible(true);
			}
		});
		btnNewButton_4.setFont(new Font("Tahoma", Font.PLAIN, 10));
		btnNewButton_4.setBounds(426, 117, 99, 23);
		frame.getContentPane().add(btnNewButton_4);
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
			JOptionPane.showMessageDialog(null,"Arquivo Criado");
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println("Arquivo não encontrado");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("erro na edição do arquivo");
		}
	}

	
	
}
