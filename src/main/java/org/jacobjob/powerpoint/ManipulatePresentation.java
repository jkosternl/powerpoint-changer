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
import java.util.Arrays;

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
        storeSlideshow(ppt, newFilename, powerPointFile.lastModified());
        ppt.close();
    }

    private void storeSlideshow(XMLSlideShow ppt, final File outputFile, long lastModified) throws IOException {
        log.info("Writing to: {}", outputFile);
        outputFile.getParentFile().mkdirs();
        outputFile.delete();
        FileOutputStream out = new FileOutputStream(outputFile);
        ppt.write(out);
        out.close();
        //Restore last modified date (from old file)
        outputFile.setLastModified(lastModified);
    }

    private void manipulateSlide(XSLFSlide slide) {
        log.info("Shapes found: {}", slide.getShapes().size());
        int shapeCounter = 0;
        final ArrayList<XSLFShape> removeList = new ArrayList<>();
        for (final XSLFShape shape : slide.getShapes()) {
            shapeCounter++;
            log.debug("Shape {} = name: {}", shapeCounter, shape.getShapeName());
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                //Remove line boxes created with text boxes, which we don't like.
                if (textShape.getText().length() == 0) {
                    removeList.add(textShape);
                    continue;
                }
                changeTextFontSize(textShape);
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
            log.debug("Removed: {}", shape.getShapeName());
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

    private void changePictureLocation(final XSLFPictureShape pictureShape) {
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
        if (textLength < 4) {
            //Signs as '>>'
            if (textLength >= 2 && textShape.getText().startsWith(">>")) {
                //Move a bit down and right, for new sized slide
                anchor.setFrame(anchor.getX() + 20, anchor.getY() + 110, anchor.getWidth(), anchor.getHeight());
            } else {
                return;
            }
        } else if (textLength < 40) {
            anchor.setFrame(newXPosition, anchor.getY(), anchor.getWidth(), anchor.getHeight());
            log.info("Changed ONLY frame position - likely title. Contents: '{}'", textShape.getText());
        } else {
            anchor.setFrame(newXPosition, newYPosition, 655d, 540d - 20d);
            log.info("Changed frame size & position");
        }
        textShape.setAnchor(anchor);
    }

    private void changeTextFontSize(final XSLFTextShape textShape) {
        final int lines = textShape.getText().split("\n").length;
        //Likely a sign (>>) or a title. Skip it.
        if (lines <= 2)
            return;
        String[] lineLengths = textShape.getText().split("\n");
        int maxLineLength = Arrays.stream(lineLengths).mapToInt(String::length).max().orElse(0);

        textShape.setTextAutofit(TextShape.TextAutofit.SHAPE);
        log.info("Found lines: {} and maxLineLength: {}", lines, maxLineLength);
        textShape.getTextParagraphs().forEach(paragraph -> {
            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                if (lines <= 12 && maxLineLength <= 33) {
                    textRun.setFontSize(37.);
                } else if (lines <= 12 && maxLineLength <= 35) {
                    textRun.setFontSize(36.);
                } else if (lines <= 12 && maxLineLength <= 36) {
                    textRun.setFontSize(34.);
                } else if (lines <= 13 && maxLineLength <= 38) {
                    textRun.setFontSize(33.);
                } else if (lines <= 14 && maxLineLength <= 46) {
                    textRun.setFontSize(29.);
                } else if (lines <= 15 && maxLineLength <= 48) {
                    textRun.setFontSize(26.);
                } else if (lines <= 16 && maxLineLength <= 55) {
                    textRun.setFontSize(25.);
                } else if (lines <= 17 && maxLineLength <= 60) {
                    textRun.setFontSize(21.);
                } else if (lines <= 19 && maxLineLength <= 65) {
                    textRun.setFontSize(17.);
                } else if (lines <= 21) {
                    textRun.setFontSize(14.);
                }
            }
        });
    }

}
