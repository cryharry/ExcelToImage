package application;

import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.ResourceBundle;

import com.extentech.ExtenXLS.ImageHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ChoiceDialog;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;
import javafx.scene.control.TextInputDialog;
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
	private String savefolder, excelFile, classStr, banStr, strNum;
	@FXML
	private ComboBox<String> classCombo;
	@FXML
	private TextField banTxt;
	
	public void doit() {
		if(excelFile == null) {
			handleOpenExcel();
		}
		
		if(savefolder == null) {
			handleSaveFolder();
		}
		
		this.classStr = classCombo.getSelectionModel().getSelectedItem();
		if(classStr == null) {
			inputClass();
		}
		
		this.banStr = banTxt.getText();
		if(banStr.equals("")) {
			intputBan();
		}
		
		WorkBookHandle tbo = new WorkBookHandle(excelFile);
		WorkSheetHandle[] sheets = tbo.getWorkSheets();
			
		for (int i = 0; i < sheets.length; i++) {
		    ImageHandle[] extracted = sheets[i].getImages();
		    for (int t = 0; t < extracted.length; t++) {
		        FileOutputStream outimg;
				try {
					if(Integer.valueOf(banStr) < 10) {
						if(Integer.valueOf(extracted[t].getName().replace("Picture ", "")) <=9){ 
							strNum = classStr+"0"+banStr+"0";
						} else {
							strNum = classStr+"0"+banStr;
						}
					} else {
						if(Integer.valueOf(extracted[t].getName().replace("Picture ", "")) <=9){ 
							strNum = classStr+banStr+"0";
						} else {
							strNum = classStr+banStr;
						}
					}
					
					outimg = new FileOutputStream(savefolder +"\\"+ strNum + extracted[t].getName().replace("Picture ", "") + ".jpg");
					extracted[t].write(outimg);
			        outimg.flush();
			        outimg.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
		        tbo.close();
		    }
		}
		Alert alert = new Alert(AlertType.INFORMATION);
        alert.setTitle("저장완료");
        alert.setHeaderText(null);
        alert.setContentText("사진파일 저장완료!");

        alert.showAndWait();
        
        excelFile = null;
        savefolder = null;

	}
	
	private void inputClass() {
		List<String> choices = new ArrayList<>();
		choices.add("1");
		choices.add("2");
		choices.add("3");

		ChoiceDialog<String> dialog = new ChoiceDialog<>("1", choices);
		dialog.setHeaderText("학년을 선택해 주세요!");
		dialog.setContentText("학년:");

		Optional<String> result = dialog.showAndWait();
		if (result.isPresent()){
		    this.classStr = result.get();
		} else {
			inputClass();
		}
	}

	private void intputBan() {
		TextInputDialog dialog = new TextInputDialog("1");
		
		dialog.setHeaderText("반을 입력해 주세요!");
		dialog.setContentText("반:");

		Optional<String> result = dialog.showAndWait();
		if(result.isPresent()) {
			this.banStr = result.get();
		} else {
			intputBan();
		}
	}

	
	
	@FXML
	public void handleOpenExcel() {
		fc = new FileChooser();
		fc.setTitle("엑셀 파일 선택 - xls");
		FileChooser.ExtensionFilter exFilter = new FileChooser.ExtensionFilter("xls file(*.xls)", "*.xls");
		fc.getExtensionFilters().add(exFilter);
		file =  fc.showOpenDialog(primaryStage);
		this.excelFile = file.getPath().replace("\\", "/");
	}
	
	@FXML
	public void handleSaveFolder() {
		DirectoryChooser dc = new DirectoryChooser();
		dc.setTitle("사진을 저장할 폴더 선택");
		file = dc.showDialog(primaryStage);
		this.savefolder = file.getPath();
	}
	
	
	@Override
	public void initialize(URL arg0, ResourceBundle arg1) {
		assert classCombo != null : "fx:id=\"classCombo\" was not injected: check your FXML file 'ExcelImage.fxml'.";
		assert banTxt != null : "fx:id=\"banTxt\" was not injected: check your FXML file 'ExcelImage.fxml'.";	
	}
}
