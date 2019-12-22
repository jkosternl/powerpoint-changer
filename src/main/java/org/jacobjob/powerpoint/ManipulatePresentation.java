package org.jacobjob.powerpoint;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ManipulatePresentation {

    private static final Logger log = LogManager.getLogger(ManipulatePresentation.class.getName());

    public void processPowerpointFile(String powerPointFile) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(powerPointFile));
        log.info("Slides found: {}", ppt.getSlides().size());

        changePowerPointResolution(ppt);

        for (XSLFSlide slide : ppt.getSlides()) {
            manipulateSlide(slide);
        }

        storeSlideshow(ppt, "output.pptx");
    }

    private void storeSlideshow(XMLSlideShow ppt, final String filename) throws IOException {
        FileOutputStream out = new FileOutputStream(filename);
        ppt.write(out);
        out.close();
        ppt.close();
    }

    private void manipulateSlide(XSLFSlide slide) {
        log.info(slide);
        log.info("Shapes found: {}", slide.getShapes().size());
        int shapeCounter = 0;
        final ArrayList<XSLFShape> removeList = new ArrayList<>();
        for (final XSLFShape shape : slide.getShapes()) {
            shapeCounter++;
            log.info("Shape {} = name: {}", shapeCounter, shape.getShapeName());
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                log.info("Contains text: {}", textShape.getText().substring(0, 40));
                changeFontSize(textShape);
                changeTextboxSize(textShape);
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
        log.info("Page size = width: {}, height: {}", pgsize.width, pgsize.height); //should be: width: 960, height: 540
        if (pgsize.width != 960) {
            ppt.setPageSize(new Dimension(960, pgsize.height));
            log.info("Changed size to 960 pixels width.");
        }
    }

    private void changeTextboxSize(final XSLFTextShape textShape) {
        final Rectangle2D anchor = textShape.getAnchor();
        anchor.setFrame(304.2324409448819d, 14.853543307086614d, 655.7675590551181d, 540d - 20d);
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
