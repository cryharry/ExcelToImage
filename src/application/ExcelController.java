package application;

import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.ResourceBundle;

import javax.imageio.ImageIO;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.ImageHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.WorkSheetNotFoundException;

import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import sun.awt.image.ByteArrayImageSource;

public class ExcelController implements Initializable {
	public static final int RATIO = 0;
    public static final int SAME = -1;
	@FXML
	private Stage primaryStage;
	@FXML
	private FileChooser fc;
	@FXML
	private File file;
	@FXML
	private String savefolder, excelFile, classStr, banStr, mapPoint;
	@FXML
	private ComboBox<String> classCombo, sheetCombo;
	@FXML
	private TextField banTxt, excelFileTxt, saveFolderTxt;
	@FXML
	private HashMap<String, String> point = new HashMap<String, String>();
	@FXML
	private WorkBookHandle work;
	@FXML
	private WorkSheetHandle sheet;
	@FXML
	private Alert alert;
	@FXML
	
	public void doit() {
		if(excelFileTxt.getText().toString().equals("")) {
			alert = new Alert(AlertType.ERROR);
	        alert.setTitle("엑셀파일이 선택되지 않았습니다.");
	        alert.setHeaderText(null);
	        alert.setContentText("엑셀파일을 선택해주세요!");
	        alert.showAndWait();
			handleOpenExcel();
		}
		if(saveFolderTxt.getText().toString().equals("")) {
			alert = new Alert(AlertType.ERROR);
	        alert.setTitle("저장할 폴더가 선택되지 않았습니다.");
	        alert.setHeaderText(null);
	        alert.setContentText("저장할 폴더를 선택해주세요!");
	        alert.showAndWait();
			handleSaveFolder();
		}
		if(sheetCombo.getSelectionModel().getSelectedItem() == null) {
			alert = new Alert(AlertType.ERROR);
	        alert.setTitle("Sheet가 선택되지 않았습니다.");
	        alert.setHeaderText(null);
	        alert.setContentText("Sheet를 선택해주세요!");
	        alert.showAndWait();
		} else {
			CellHandle[] cellHandle = work.getCells();
			for(int i=0;i<cellHandle.length;i++) {
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
			ImageHandle[] extracted  = sheet.getImages();
			for(int k=0; k<extracted.length; k++) {
					int l= extracted[k].getRow()+2;
					int m = extracted[k].getCol()+1;
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
						String getPoint = "";
						if(point.get(mapPoint) != null) {
							if(Integer.valueOf(point.get(mapPoint))<10) {
								getPoint = "0" +point.get(mapPoint);
							} else {
								getPoint = point.get(mapPoint);
							}
						}
						if(getPoint != "") {
							FileOutputStream outimg;
							try {
								String fileName = savefolder+"\\"+classStr+banStr+getPoint+".jpg";
								outimg = new FileOutputStream(fileName);
								byte[] bytes = extracted[k].getImageBytes();
								BufferedImage image = ImageIO.read(new ByteArrayInputStream(bytes));
								resize(image, 300, 400);
								//extracted[k].write(outimg);
								//outimg.flush();
								//outimg.close();
							} catch (Exception e) {
								e.printStackTrace();
							}
						}
			}
			alert = new Alert(AlertType.INFORMATION);
	        alert.setTitle("저장완료");
	        alert.setHeaderText(null);
	        alert.setContentText("사진파일 저장완료!");
	        alert.showAndWait();
		}
	}
	
	public void resize(BufferedImage src, int maxWidth, int maxHeight) throws IOException {
		//BufferedImage src = ImageIO.read(Image);

		try {
			int width = src.getWidth();
			int height = src.getHeight();
	 
			if (width > maxWidth) {
				float widthRatio = maxWidth / (float) width;
				width = (int) (width * widthRatio);
	            height = (int) (height * widthRatio);
	        }
	        if (height > maxHeight) {
	        	float heightRatio = maxHeight / (float) height;
	            width = (int) (width * heightRatio);
	            height = (int) (height * heightRatio);
	        }
	        BufferedImage destImg = new BufferedImage(width, height, BufferedImage.TYPE_3BYTE_BGR);
	        Graphics2D g = destImg.createGraphics();
	        g.drawImage(src, 0, 0, width, height, null);
	        //ByteArrayOutputStream bo = new ByteArrayOutputStream();
	        FileOutputStream fos = new FileOutputStream(savefolder);
	        ImageIO.write(destImg, "jpg", fos);
	        fos.flush();
	        fos.close();
	        //return destImg;
	    } catch (Exception e) {
	    	e.printStackTrace();
	    	//System.out.println("이 파일은 사이즈를 변경할 수 없습니다. 파일을 확인해주세요.");
	    	//return null;
	    }
	}
	
	public ObservableList<String> getSheetsName() {
		ObservableList<String> comboList = FXCollections.observableArrayList();
		WorkSheetHandle[] workSheets = work.getWorkSheets();
		for(int i=0;i<workSheets.length;i++) {
			comboList.add(workSheets[i].getSheetName());
		}
		return comboList;
	}
	
	@FXML
	public void handleOpenExcel() {
		fc = new FileChooser();
		fc.setTitle("엑셀 파일 선택 - xls");
		FileChooser.ExtensionFilter xlsFilter = new FileChooser.ExtensionFilter("xls file(*.xls,*xlsx)", "*.xls","*.xlsx");
		fc.getExtensionFilters().add(xlsFilter);
		file =  fc.showOpenDialog(primaryStage);
		if(file == null) {
			handleOpenExcel();
		}
		this.excelFile = file.getPath().replace("\\", "/");
		excelFileTxt.setText(excelFile);
		work = new WorkBookHandle(excelFile);
		ObservableList<String> sheetName = getSheetsName();
		sheetCombo.setItems(sheetName);
		sheetCombo.valueProperty().addListener(new ChangeListener<String>() {
			@Override
			public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
				try {
					sheet = work.getWorkSheet(newValue);
				} catch (WorkSheetNotFoundException e) {
					e.printStackTrace();
				}
			}
		});
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
		saveFolderTxt.setText(savefolder);
	}
	
	@Override
	public void initialize(URL arg0, ResourceBundle arg1) {
	}
}
