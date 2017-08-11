package application;

import java.awt.Graphics2D;
import java.awt.RenderingHints;
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
import com.mortennobel.imagescaling.ResampleOp;
import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGEncodeParam;
import com.sun.image.codec.jpeg.JPEGImageEncoder;

import javafx.beans.value.ChangeListener;
import javafx.beans.value.ObservableValue;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
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
	private String savefolder, excelFile, classStr, banStr, mapPoint, selSheet = "";
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
	private Button excelBtn, folderBtn, saveBtn;
	
	public void doit() {
		if(excelFileTxt.getText().toString().equals("")) {
			alert = new Alert(AlertType.ERROR);
	        alert.setTitle("엑셀파일이 선택되지 않았습니다.");
	        alert.setHeaderText(null);
	        alert.setContentText("엑셀파일을 선택해주세요!");
	        alert.showAndWait();
			handleOpenExcel();
		}
		if(sheetCombo.getSelectionModel().getSelectedItem() == null) {
			alert = new Alert(AlertType.ERROR);
	        alert.setTitle("Sheet가 선택되지 않았습니다.");
	        alert.setHeaderText(null);
	        alert.setContentText("Sheet를 선택해주세요!");
	        alert.showAndWait();
		}
		try {
			sheet = work.getWorkSheet(selSheet);
		} catch (WorkSheetNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		if(saveFolderTxt.getText().toString().equals("")) {
			alert = new Alert(AlertType.ERROR);
	        alert.setTitle("저장할 폴더가 선택되지 않았습니다.");
	        alert.setHeaderText(null);
	        alert.setContentText("저장할 폴더를 선택해주세요!");
	        alert.showAndWait();
			handleSaveFolder();
		}
			CellHandle[] cellHandle = sheet.getCells();
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
								outimg.write(resize(image, 300, 1.3));
								//outimg.write(reSample(image, 300, 400, 300));
								//extracted[k].write(outimg);
								outimg.flush();
								outimg.close();
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
	
	public static byte[] resize(BufferedImage src, int maxWidth, double xyRatio) throws IOException {
		
		//이미지 자를 포지션 설정
		int[] centerPoint = {src.getWidth() / 2, src.getHeight() / 2};
		
		//받은 이미지 사이즈
		int cropWidth = src.getWidth();
		int cropHeight = src.getHeight();
	    
		if (cropHeight > cropWidth * xyRatio) {
			cropHeight = (int) ( cropWidth * xyRatio );
	    } else {
	       	cropWidth = (int) ( (float) cropHeight / xyRatio );
	    }
		
		//저장될 이미지 사이즈
		int targetWidth = cropWidth;
		int targetHeight = cropHeight;
			
		if(targetWidth > maxWidth) {
			targetWidth = maxWidth;
			targetHeight = (int)(targetWidth * xyRatio);
		}
	       
	    BufferedImage destImg = new BufferedImage(targetWidth, targetHeight, BufferedImage.TYPE_INT_RGB);
	    Graphics2D g = destImg.createGraphics();
	    g.setRenderingHint(RenderingHints.KEY_INTERPOLATION,RenderingHints.VALUE_INTERPOLATION_BILINEAR);
	    g.setRenderingHint(RenderingHints.KEY_RENDERING,RenderingHints.VALUE_RENDER_QUALITY);
	    g.setRenderingHint(RenderingHints.KEY_ANTIALIASING,RenderingHints.VALUE_ANTIALIAS_ON);
	    g.drawImage(src, 0, 0, targetWidth, targetHeight, centerPoint[0] - (int)(cropWidth /2) , centerPoint[1] - (int)(cropHeight /2), centerPoint[0] + (int)(cropWidth /2), centerPoint[1] + (int)(cropHeight /2), null);
	    
	    byte[] returnByte = reSample(destImg, 300, 400, 300); 
	    return returnByte;
	}
	public static byte[] reSample (BufferedImage src, int width, int height, int dpi) throws IOException {
		ResampleOp resampleOp = new ResampleOp(width,height);
	    BufferedImage bsImg = resampleOp.filter(src, null);
	    
	    ByteArrayOutputStream bos = new ByteArrayOutputStream();
	    JPEGImageEncoder jpgEncoder = JPEGCodec.createJPEGEncoder(bos);
	    JPEGEncodeParam jpgParam = jpgEncoder.getDefaultJPEGEncodeParam(bsImg);
	    jpgParam.setDensityUnit(JPEGEncodeParam.DENSITY_UNIT_DOTS_INCH);
	    jpgEncoder.setJPEGEncodeParam(jpgParam);
	    jpgParam.setQuality(1, false);
	    jpgParam.setXDensity(dpi); jpgParam.setYDensity(dpi);
	    jpgEncoder.encode(bsImg, jpgParam);
	    
	    ImageIO.write(bsImg, "jpg", bos);
	    
	    return bos.toByteArray();
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
		FileChooser.ExtensionFilter xlsFilter = new FileChooser.ExtensionFilter("xls file(*.xls)", "*.xls");
		fc.getExtensionFilters().add(xlsFilter);
		file =  fc.showOpenDialog(primaryStage);
		if(file == null) {
			handleOpenExcel();
		}
		this.excelFile = file.getPath().replace("\\", "/");
		excelFileTxt.setText(excelFile);
		work = new WorkBookHandle(file);
		ObservableList<String> sheetName = getSheetsName();
		sheetCombo.setItems(sheetName);
		sheetCombo.valueProperty().addListener(new ChangeListener<String>() {
			@Override
			public void changed(ObservableValue<? extends String> observable, String oldValue, String newValue) {
				selSheet = newValue;
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
