import Excel.ImportExportExcel;
import model.CadNumber;
import model.XlsxContainer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import parsers.FondCadCost;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.io.IOException;
import java.util.List;
import java.util.Set;
import javax.swing.*;
import javax.swing.border.EmptyBorder;

public class MainFrame extends JFrame {

	private JPanel contentPane;
	private JTextField numberSheetField;
	private JTextField cadNumberField;
	private JLabel costLabel;
	private JTextField costNumberFiled;
	private JLabel areaLabel;
	private JTextField areaNumberField;
	private JLabel selectLabel;
	private JRadioButton fondRadioButton;
	private JRadioButton pkkRadioButton;
	private JProgressBar progressBar1;
	private JButton downloadExcelButton;
	private JButton updateCostButton;
	private JTextArea textInfoArea;
	private JTextArea textErrorInfo;
	private JFileChooser fileChooser;
	private ImportExportExcel importExportExcel;
	private static List<CadNumber> listCad;
	private XlsxContainer xlsxContainer;
	private Set<String> cadSet;

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainFrame frame = new MainFrame();
					frame.setVisible(true);
					frame.showEvent();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public MainFrame() {
		importExportExcel = new ImportExportExcel();
		xlsxContainer = new XlsxContainer();

		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 600, 350);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));

		setContentPane(contentPane);
		contentPane.setLayout(null);

		JLabel labelNumberSheet = new JLabel("Номер листа");
		labelNumberSheet.setBounds(5, 6, 86, 19);
		labelNumberSheet.setHorizontalAlignment(SwingConstants.LEFT);
		labelNumberSheet.setVerticalAlignment(SwingConstants.TOP);
		contentPane.add(labelNumberSheet);

		numberSheetField = new JTextField();
		numberSheetField.setBounds(5, 35, 185, 19);
		contentPane.add(numberSheetField);
		numberSheetField.setColumns(10);

		JLabel cadLabel = new JLabel("Номер столбца с кад. номером");
		cadLabel.setVerticalAlignment(SwingConstants.TOP);
		cadLabel.setHorizontalAlignment(SwingConstants.LEFT);
		cadLabel.setBounds(5, 64, 200, 19);
		contentPane.add(cadLabel);

		cadNumberField = new JTextField();
		cadNumberField.setColumns(10);
		cadNumberField.setBounds(5, 93, 185, 19);
		contentPane.add(cadNumberField);

		costLabel = new JLabel("Номер столбца со стоимостью");
		costLabel.setVerticalAlignment(SwingConstants.TOP);
		costLabel.setHorizontalAlignment(SwingConstants.LEFT);
		costLabel.setBounds(5, 131, 200, 19);
		contentPane.add(costLabel);

		costNumberFiled = new JTextField();
		costNumberFiled.setColumns(10);
		costNumberFiled.setBounds(5, 160, 185, 19);
		contentPane.add(costNumberFiled);

		areaLabel = new JLabel("Номер столбца с площадью");
		areaLabel.setVerticalAlignment(SwingConstants.TOP);
		areaLabel.setHorizontalAlignment(SwingConstants.LEFT);
		areaLabel.setBounds(5, 189, 200, 19);
		contentPane.add(areaLabel);

		areaNumberField = new JTextField();
		areaNumberField.setColumns(10);
		areaNumberField.setBounds(5, 218, 185, 19);
		contentPane.add(areaNumberField);

		selectLabel = new JLabel("Выберите источник обновления:");
		selectLabel.setVerticalAlignment(SwingConstants.TOP);
		selectLabel.setHorizontalAlignment(SwingConstants.LEFT);
		selectLabel.setBounds(211, 6, 365, 19);
		contentPane.add(selectLabel);

		fondRadioButton = new JRadioButton("Фонд кадастровой оценки (стоимость)");
		fondRadioButton.setBounds(211, 34, 365, 21);
		fondRadioButton.setEnabled(false);
		contentPane.add(fondRadioButton);

		pkkRadioButton = new JRadioButton("Публичная кадастровая карта (стоимость и площадь)");
		pkkRadioButton.setBounds(211, 64, 365, 21);
		pkkRadioButton.setEnabled(false);
		contentPane.add(pkkRadioButton);

		progressBar1 = new JProgressBar();
		progressBar1.setBounds(211, 168, 365, 11);
		contentPane.add(progressBar1);

		downloadExcelButton = new JButton("Загрузка Excel");
		downloadExcelButton.setBounds(211, 217, 127, 21);
		downloadExcelButton.setEnabled(false);
		contentPane.add(downloadExcelButton);

		updateCostButton = new JButton("Обновить информацию");
		updateCostButton.setBounds(348, 217, 228, 21);
		updateCostButton.setEnabled(false);
		contentPane.add(updateCostButton);

		textInfoArea = new JTextArea();
		textInfoArea.setBackground(new Color(240, 240, 240));
		textInfoArea.setEditable(false);
		textInfoArea.setBounds(211, 95, 365, 63);
		textInfoArea.setText("Введите параметры столбцов файла и нажмите \"Enter\" выберите\nисточник обновления загрузите файл и нажмите\n\"обновить информацию\".");
		contentPane.add(textInfoArea);

		textErrorInfo = new JTextArea();
		textErrorInfo.setBackground(new Color(240, 240, 240));
		textErrorInfo.setForeground(new Color(238, 0, 0));
		textErrorInfo.setEditable(false);
		textErrorInfo.setBounds(10, 257, 566, 46);
		contentPane.add(textErrorInfo);
	}

	private void showEvent() {
		updateCostButton.addActionListener(new ButtonUpdateClickListener());
		downloadExcelButton.addActionListener(new ButtonDownloadClickListener());
		fondRadioButton.addActionListener(new RadioButtonClickListener());
		numberSheetField.addKeyListener(new JTextActionListener());
		cadNumberField.addKeyListener(new JTextActionListener());
		costNumberFiled.addKeyListener(new JTextActionListener());
		areaNumberField.addKeyListener(new JTextActionListener());
		progressBar1.setVisible(true);
		progressBar1.setMinimum(0);
		progressBar1.setMaximum(100);
	}

	private class ButtonDownloadClickListener implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			String[][] xlsxFilter = {{".xlsx", "Файлы"}};
			fileChooser = new JFileChooser();
			int ret = fileChooser.showDialog(null, "Открыть файл");
			if (ret == JFileChooser.APPROVE_OPTION) {

				try {
					XlsxContainer container = importExportExcel.getXSSFSheetAndWorkbook(fileChooser.getSelectedFile(), xlsxContainer.getIndexSheetRow());
					xlsxContainer.setWorkbook(container.getWorkbook());
					xlsxContainer.setSheet(container.getSheet());
					xlsxContainer.setNameFile(container.getNameFile());
					cadSet = importExportExcel.getSetCad(xlsxContainer);
				}catch (NotOfficeXmlFileException ex){
					textErrorInfo.setText("ОШИБКА: выбран неправильный формат файла.");
					updateCostButton.setEnabled(false);
				}catch (IllegalArgumentException ex){
					textErrorInfo.setText("ОШИБКА: некорректный файл Excel.");
					updateCostButton.setEnabled(false);
				}
				catch (IOException ex) {
					textErrorInfo.setText("ОШИБКА: ввода ввыода - "+ex.getMessage());
					updateCostButton.setEnabled(false);
				} catch (InvalidFormatException ex) {
					textErrorInfo.setText("ОШИБКА: "+ ex.getMessage());
					updateCostButton.setEnabled(false);
				}
			}
			downloadExcelButton.setText("Загружено");
			updateCostButton.setEnabled(true);
		}
	}

	private class RadioButtonClickListener implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			//&& cadSet != null
			if (fondRadioButton.isSelected()) {
				downloadExcelButton.setEnabled(true);
			} else if (fondRadioButton.isSelected() && cadSet != null) {
				updateCostButton.setEnabled(true);
			} else {
				updateCostButton.setEnabled(false);
			}
		}
	}

	private class JTextActionListener implements KeyListener {
		@Override
		public void keyTyped(KeyEvent e) {
		}
		@Override
		public void keyPressed(KeyEvent e) {
			System.out.println(numberSheetField.getText()+" -- "+cadNumberField.getText() + " -- "+costNumberFiled.getText());
			progressBar1.setValue(0);
			if (numberSheetField.getText().matches("\\d+") && cadNumberField.getText().matches("\\d+")
					&& costNumberFiled.getText().matches("\\d+") && areaNumberField.getText().matches("\\d+")) {
				textErrorInfo.setText("");
				xlsxContainer.setIndexSheetRow(Integer.parseInt(numberSheetField.getText().trim()));
				xlsxContainer.setCadRow(Integer.parseInt(cadNumberField.getText().trim()));
				xlsxContainer.setCostRow(Integer.parseInt(costNumberFiled.getText().trim()));
				xlsxContainer.setAreaRow(Integer.parseInt(areaNumberField.getText().trim()));
				fondRadioButton.setEnabled(true);
				pkkRadioButton.setEnabled(true);
			} else {
				if(e.getKeyCode() == KeyEvent.VK_ENTER) {
					textErrorInfo.setText("ОШИБКА: введено текстовое значения в полях!");
					fondRadioButton.setEnabled(false);
					pkkRadioButton.setEnabled(false);
					downloadExcelButton.setEnabled(false);
				}
			}
		}
		@Override
		public void keyReleased(KeyEvent e) {
		}
	}

	private class ButtonUpdateClickListener implements ActionListener {
		@Override
		public void actionPerformed(ActionEvent e) {
			progressBar1.setEnabled(true);
			progressBar1.setVisible(true);
			updateCostButton.setEnabled(false);
			downloadExcelButton.setText("Загрузка Excel");
			FondCadCost cadCost = new FondCadCost();
			WebDriver driver = new FirefoxDriver();
			if (cadSet != null) {
				if (fondRadioButton.isSelected()) {

					progressBar1.setIndeterminate(true);
					progressBar1.setValue(FondCadCost.getIndexProgress());
					new Thread(new Runnable() {
						@Override
						public void run() {
							try {
								listCad = cadCost.getListCad(cadSet, driver);
								importExportExcel.updateCostToExcel(xlsxContainer, listCad);
								progressBar1.setIndeterminate(false);
								progressBar1.setStringPainted(true);
								progressBar1.setValue(100);
							} catch (InvalidFormatException ex) {
								ex.printStackTrace();
							} catch (IOException ex) {
								ex.printStackTrace();
							}
						}
					}).start();

				} else if (pkkRadioButton.isSelected()) {

				} else {
					updateCostButton.setEnabled(false);
				}
			}
			updateCostButton.setEnabled(false);
			fondRadioButton.setSelected(false);
			fondRadioButton.setEnabled(false);
			pkkRadioButton.setEnabled(false);
			numberSheetField.setText("");
			cadNumberField.setText("");
			costNumberFiled.setText("");
			areaNumberField.setText("");
			downloadExcelButton.setEnabled(false);
		}
	}
}
