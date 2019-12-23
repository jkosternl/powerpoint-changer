package org.jacobjob.powerpoint;

import lombok.extern.log4j.Log4j2;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

@Log4j2
public class ManipulatePresentation {

    public static final double newXPosition = 304.2324409448819d;
    public static final double newYPosition = 14.853543307086614d;

    public void processPowerpointFile(File powerPointFile) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(powerPointFile));
        log.info("Starting with file: {} - Slides found: {}", powerPointFile.getAbsolutePath(), ppt.getSlides().size());

        changePowerPointResolution(ppt);

        for (XSLFSlide slide : ppt.getSlides()) {
            manipulateSlide(slide);
        }

//        storeSlideshow(ppt, "output.pptx");
    }

    private void storeSlideshow(XMLSlideShow ppt, final String filename) throws IOException {
        FileOutputStream out = new FileOutputStream(filename);
        ppt.write(out);
        out.close();
        ppt.close();
    }

    private void manipulateSlide(XSLFSlide slide) {
        log.info("Shapes found: {}", slide.getShapes().size());
        int shapeCounter = 0;
        final ArrayList<XSLFShape> removeList = new ArrayList<>();
        for (final XSLFShape shape : slide.getShapes()) {
            shapeCounter++;
            log.info("Shape {} = name: {}", shapeCounter, shape.getShapeName());
            if (shapeCounter > 2) log.warn("Shape = {} - {}", shapeCounter, shape);
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                changeFontSize(textShape);
                changeTextboxSize(textShape);
            }
            if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
                changePictureLocation(pictureShape);
                changePictureSize(pictureShape);
            }
            if (shape instanceof XSLFGroupShape) {
                XSLFGroupShape groupShape = (XSLFGroupShape) shape;
                removeList.add(groupShape);
            }
        }
        for (final XSLFShape shape : removeList) {
            log.info("Removed: {}", shape.getShapeName());
            slide.removeShape(shape);
        }
    }

    private void changePowerPointResolution(final XMLSlideShow ppt) {
        Dimension pgsize = ppt.getPageSize();
        log.debug("Page size = width: {}, height: {}", pgsize.width, pgsize.height); //should be: width: 960, height: 540
        if (pgsize.width != 960) {
            ppt.setPageSize(new Dimension(960, pgsize.height));
            log.debug("Changed size to 960 pixels width.");
        }
    }

    private void changePictureSize(XSLFPictureShape pictureShape) {
        final Rectangle2D anchor = pictureShape.getAnchor();
        final double ratio = anchor.getHeight() / anchor.getWidth();
    }

    private void changePictureLocation(XSLFPictureShape pictureShape) {
        final Rectangle2D anchor = pictureShape.getAnchor();
        // If this is a full screen picture, do not move it to the right.
        if (anchor.getX() < newXPosition) {
            return;
        }
        double height = anchor.getHeight();
        double width = anchor.getWidth();
        anchor.setFrame(newXPosition, anchor.getY(), height, width);
        pictureShape.setAnchor(anchor);
        log.info("Picture adjusted");
    }

    private void changeTextboxSize(final XSLFTextShape textShape) {
        //Skip only titles
        int textLength = textShape.getText().length();
        final Rectangle2D anchor = textShape.getAnchor();
        if (textLength < 40) {
            anchor.setFrame(newXPosition, anchor.getY(), anchor.getHeight(), anchor.getWidth());
            log.info("Changed ONLY frame position - likely title.");
        } else {
            anchor.setFrame(newXPosition, newYPosition, 655.7675590551181d, 540d - 20d);
            log.info("Changed frame size & position");
        }
        textShape.setAnchor(anchor);
    }

    private void changeFontSize(final XSLFTextShape textShape) {
        textShape.setTextAutofit(TextShape.TextAutofit.SHAPE);
        textShape.getTextParagraphs().forEach(paragraph -> {
            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                textRun.setFontSize(28.);
            }
        });
    }

}
