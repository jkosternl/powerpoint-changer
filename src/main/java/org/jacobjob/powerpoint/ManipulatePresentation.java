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

    public static final double newXPosition = 304d;
    public static final double newYPosition = 15d;

    public void processPowerpointFile(File powerPointFile) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(powerPointFile));
        log.info("Starting with file: {} - Slides found: {}", powerPointFile.getAbsolutePath(), ppt.getSlides().size());

        changePowerPointResolution(ppt);

        for (XSLFSlide slide : ppt.getSlides()) {
            manipulateSlide(slide);
        }

        File newFilename = new File(powerPointFile.getAbsolutePath().replaceFirst("Databank", "Databank-output"));
        storeSlideshow(ppt, newFilename);
        ppt.close();
    }

    private void storeSlideshow(XMLSlideShow ppt, final File outputFile) throws IOException {
        log.info("Writing to: {}", outputFile);
        outputFile.getParentFile().mkdirs();
        outputFile.delete();
        FileOutputStream out = new FileOutputStream(outputFile);
        ppt.write(out);
        out.close();
    }

    private void manipulateSlide(XSLFSlide slide) {
        log.info("Shapes found: {}", slide.getShapes().size());
        int shapeCounter = 0;
        final ArrayList<XSLFShape> removeList = new ArrayList<>();
        for (final XSLFShape shape : slide.getShapes()) {
            shapeCounter++;
            log.info("Shape {} = name: {}", shapeCounter, shape.getShapeName());
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                //Remove line boxes created with text boxes, which we don't like.
                if (textShape.getText().length() == 0) {
                    removeList.add(textShape);
                    continue;
                }
                changeFontSize(textShape);
                changeTextboxSize(textShape);
            } else if (shape instanceof XSLFPictureShape) {
                XSLFPictureShape pictureShape = (XSLFPictureShape) shape;
                changePictureLocation(pictureShape);
                pictureShape.setAnchor(changePictureAnchorSize(pictureShape.getAnchor()));
            } else if (shape instanceof XSLFGroupShape) {
                XSLFGroupShape groupShape = (XSLFGroupShape) shape;
                removeList.add(groupShape);
            } else {
                log.warn("Not recognized shape: {}", shape.getShapeName());
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

    Rectangle2D changePictureAnchorSize(final Rectangle2D anchor) {
        final double ratio = anchor.getHeight() / anchor.getWidth();
        double newWidth = 960 - newXPosition - 5;
        double newHeight = newWidth * ratio;
        //New width is too large, start again calculations on width.
        if (newHeight > (540 - newYPosition)) {
            newHeight = (540 - newYPosition - 5);
            newWidth = newHeight / ratio;
        }
        log.info("Picture resized from (H x W): {} x {} to: {} x {}", anchor.getHeight(), anchor.getWidth(), newHeight, newWidth);
        anchor.setFrame(anchor.getX(), anchor.getY(), Math.floor(newWidth), Math.floor(newHeight));
        return anchor;
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
            anchor.setFrame(newXPosition, anchor.getY(), anchor.getWidth(), anchor.getHeight());
            log.info("Changed ONLY frame position - likely title. Contents: '{}'", textShape.getText());
        } else {
            anchor.setFrame(newXPosition, newYPosition, 655d, 540d - 20d);
            log.info("Changed frame size & position");
        }
        textShape.setAnchor(anchor);
    }

    private void changeFontSize(final XSLFTextShape textShape) {
        final int lines = textShape.getText().split("\n").length;
        textShape.setTextAutofit(TextShape.TextAutofit.SHAPE);
        textShape.getTextParagraphs().forEach(paragraph -> {
            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                if (lines <= 15) {
                    textRun.setFontSize(28.);
                } else if (lines <= 17) {
                    textRun.setFontSize(24.);
                } else if (lines <= 19) {
                    textRun.setFontSize(20.);
                } else if (lines <= 21) {
                    textRun.setFontSize(16.);
                }
            }
        });
    }

}
