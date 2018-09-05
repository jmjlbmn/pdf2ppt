package vip.maning.pdf2ppt;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.Tooltip;
import javafx.scene.image.Image;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.stage.Stage;
import net.coobird.thumbnailator.Thumbnails;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * @author Maning
 */
public class App extends Application {
    private Rectangle rectangle;
    private ProgressBar progress;
    private boolean start = false;
    private Stage stage;
    private ExecutorService pool = Executors.newSingleThreadExecutor();

    public static void main(String[] args) {
        App.launch(args);
    }

    @Override
    public void start(Stage stage) throws Exception {
        Text text = new Text("拖放文件到此处");
        text.setFont(new Font(24));
        rectangle = new Rectangle(450, 180, Color.TRANSPARENT);
        rectangle.setStroke(Color.rgb(238, 130, 238));
        progress = new ProgressBar();
        progress.setMaxWidth(400);
        progress.setVisible(false);
        this.stage = stage;
        VBox root = new VBox();
        root.setPadding(new Insets(10));
        root.setSpacing(10);
        root.setAlignment(Pos.CENTER);
        root.getChildren().add(text);
        rectangle.setOnDragOver(event -> {
            if (!start) {
                event.acceptTransferModes(TransferMode.ANY);
                rectangle.setFill(Color.rgb(255, 0, 0, 0.4));
            }
        });
        rectangle.setOnDragExited(event -> {
            rectangle.setFill(Color.TRANSPARENT);
        });
        rectangle.setOnDragDropped(event -> {
            if (!start) {
                Dragboard dragboard = event.getDragboard();
                if (dragboard.hasFiles()) {
                    File file = dragboard.getFiles().get(0);
                    if (file.getName().toLowerCase().endsWith("pdf")) {
                        turnStart();
                        doTurn(file);
                    }
                }
            }
        });
        root.getChildren().add(rectangle);
        root.getChildren().add(progress);
        Scene scene = new Scene(root, 521, 258);
        stage.setOpacity(0.8);
        stage.setMaximized(false);
        stage.getIcons().add(new Image(getClass().getResourceAsStream("/icon.png")));
        stage.setTitle("PDF to PPT");
        stage.setScene(scene);
        stage.setResizable(false);
        stage.show();
        stage.setOnCloseRequest(e -> pool.shutdownNow());
    }

    private void turnStart() {
        this.start = true;
        progress.setVisible(true);
        progress.setProgress(0);
    }

    private void turnStop() {
        this.start = false;
        progress.setVisible(false);
    }

    private void doTurn(File pdfFile) {
        pool.submit(() -> {
            try {
                PDDocument document = PDDocument.load(pdfFile);
                PDFRenderer render = new PDFRenderer(document);
                XMLSlideShow ppt = new XMLSlideShow();
                int count = document.getNumberOfPages();
                for (int i = 0; i < count; i++) {
                    double jd = (i + 1) / (double) count;
                    Platform.runLater(() -> {
                        progress.setProgress(jd);
                    });
                    ByteArrayOutputStream bs = new ByteArrayOutputStream();
                    Thumbnails.of(render.renderImage(i, 3, ImageType.ARGB)).size(1440, 1080).outputQuality(1).outputFormat("png").toOutputStream(bs);
                    XSLFPictureShape shape = ppt.createSlide().createPicture(ppt.addPicture(new ByteArrayInputStream(bs.toByteArray()), PictureData.PictureType.PNG));
                    shape.setAnchor(new java.awt.Rectangle(0, 0, 720, 540));
                }
                ppt.write(new FileOutputStream(pdfFile.getParent() + File.separatorChar + pdfFile.getName() + "转换的PPT文件.pptx"));
                Platform.runLater(this::turnStop);
            } catch (Exception e) {
                e.printStackTrace();
                Platform.runLater(() -> {
                    turnStop();
                    Tooltip tooltip = new Tooltip("转换失败" + e.getMessage());
                    tooltip.show(stage);
                });
            }
        });
    }
}
