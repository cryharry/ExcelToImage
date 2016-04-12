package application;

import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.util.HashMap;
import java.util.ResourceBundle;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.ImageHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class ExcelController implements Initializable {
	@FXML
	private Stage primaryStage;
	@FXML
	private FileChooser fc;
	@FXML
	private File file;
	@FXML
	private String savefolder, excelFile, classStr, banStr, mapPoint;
	@FXML
	private ComboBox<String> classCombo;
	@FXML
	private TextField banTxt;
	@FXML
	private HashMap<String, String> point = new HashMap<String, String>();
	
	public void doit() {
		if(excelFile == null) {
			handleOpenExcel();
		}
		if(savefolder == null) {
			handleSaveFolder();
		}
		
		WorkBookHandle tbo = new WorkBookHandle(excelFile);
		CellHandle[] cellHandle = tbo.getCells();
		for(int i=3;i<cellHandle.length;i++) {
				if(cellHandle[i].getCell().toString().startsWith("LABELSST")) {
					String reCell = cellHandle[i].getCell().toString().replace("LABELSST:", "");
					if(i<5 && reCell.contains("학년")) {
						classStr = reCell.split(":")[1].substring(reCell.split(":")[1].indexOf("학년도")+5,reCell.split(":")[1].length()).substring(0, 1);
						if(reCell.contains("반")){
							banStr = reCell.split(":")[1].substring(reCell.split(":")[1].lastIndexOf("반")-1,reCell.split(":")[1].length()-1);
						}
					}
					if(reCell.contains("번")){
						mapPoint = reCell.split(":")[0];
						point.put(mapPoint,reCell.split(":")[1].substring(0, reCell.split(":")[1].indexOf("번")));
					}
				}
		}
		
		if(Integer.valueOf(banStr)<10) {
			banStr = "0" + banStr;
		}
		WorkSheetHandle[] sheets = tbo.getWorkSheets();
		try {
			for(int j=0;j<sheets.length;j++){
			ImageHandle[] extracted  = sheets[j].getImages();
				for(int k=0; k<extracted.length; k++) {
					int l= extracted[k].getRow()+2;
					
					int m = extracted[k].getCol()+1;
					
					System.out.println(m+"/"+l+":"+extracted[k].getName());
					String x = "";
					switch (m) {
					case 1:
						x = "A";
						break;
					case 2:
						x = "B";
						break;
					case 3:
						x = "C";
						break;
					case 4:
						x = "D";
						break;
					case 5:
						x = "E";
						break;
					case 6:
						x = "F";
						break;
					case 7:
						x = "G";
						break;
					case 8:
						x = "H";
						break;
					case 9:
						x = "I";
						break;
					case 10:
						x = "J";
						break;
					case 11:
						x = "K";
					case 12:
						x = "L";
						break;
					case 13:
						x = "M";
						break;
					case 14:
						x = "L";
						break;
					case 15:
						x = "M";
						break;
					case 16:
						x = "N";
						break;
					case 17:
						x = "Q";
						break;
					case 18:
						x = "R";
						break;
					case 19:
						x = "S";
						break;
					case 20:
						x = "T";
						break;
					case 21:
						x = "U";
						break;
					case 22:
						x = "V";
						break;
					case 23:
						x = "W";
						break;
					case 24:
						x = "X";
						break;
					case 25:
						x = "Y";
						break;
					case 26:
						x = "Z";
						break;
					default:
						break;
					}
					mapPoint = x+l;
					System.out.println(point.get(mapPoint));
					String getPoint = "";
					if(point.get(mapPoint) != null) {
						if(Integer.valueOf(point.get(mapPoint))<10) {
							getPoint = "0" +point.get(mapPoint);
						} else {
							getPoint = point.get(mapPoint);
						}
					}
					FileOutputStream outimg = new FileOutputStream(savefolder+"\\"+classStr+banStr+getPoint+".jpg");
					extracted[k].write(outimg);
					outimg.flush();
					outimg.close();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		tbo.close();
		
		Alert alert = new Alert(AlertType.INFORMATION);
        alert.setTitle("저장완료");
        alert.setHeaderText(null);
        alert.setContentText("사진파일 저장완료!");

        alert.showAndWait();
        
        excelFile = null;
        savefolder = null;

	}

	@FXML
	public void handleOpenExcel() {
		fc = new FileChooser();
		fc.setTitle("엑셀 파일 선택 - xls");
		FileChooser.ExtensionFilter xlsFilter = new FileChooser.ExtensionFilter("xls file(*.xls)", "*.xls");
		fc.getExtensionFilters().add(xlsFilter);
		file =  fc.showOpenDialog(primaryStage);
		if(file == null) {
			handleOpenExcel();
		}
		this.excelFile = file.getPath().replace("\\", "/");
	}
	
	@FXML
	public void handleSaveFolder() {
		DirectoryChooser dc = new DirectoryChooser();
		dc.setTitle("사진을 저장할 폴더 선택");
		file = dc.showDialog(primaryStage);
		if(file == null) {
			handleSaveFolder();
		}
		this.savefolder = file.getPath();
	}
	
	
	@Override
	public void initialize(URL arg0, ResourceBundle arg1) {
	}
}
