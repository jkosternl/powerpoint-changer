package org.jacobjob.powerpoint;

import org.apache.commons.lang3.time.StopWatch;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class PowerPointChanger {

    private static final String POWERPOINT_FILE = "Opw 42.pptx";
    private static final String POWERPOINT_FILE2 = "Opw 585.pptx";

    private static final Logger log = LogManager.getLogger(PowerPointChanger.class.getName());

    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        log.info("Starting");
        memStats();

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(POWERPOINT_FILE));

        log.info("Slides found: {}", ppt.getSlides().size());
        Dimension pgsize = ppt.getPageSize();
        log.info("Page size = width: {}, height: {}", pgsize.width, pgsize.height); //should be: width: 960, height: 540
        if (pgsize.width != 960) {
            ppt.setPageSize(new Dimension(960, pgsize.height));
            log.info("Changed size to 960 pixels width.");
        }

        for (XSLFSlide slide : ppt.getSlides()) {
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
            for (final XSLFShape shape : removeList){
                log.info("Removed: {}", shape.getShapeName());
                slide.removeShape(shape);
            }
        }

        //Store
        FileOutputStream out = new FileOutputStream("output.pptx");
        ppt.write(out);
        out.close();
        ppt.close();


        memStats();
        stopWatch.stop();
        log.info("Programma beeindigd na: {}", stopWatch);
    }

    private static void changeTextboxSize(final XSLFTextShape textShape) {
        final Rectangle2D anchor = textShape.getAnchor();
        anchor.setFrame(304.2324409448819d,14.853543307086614d,655.7675590551181d, 540d-20d);
        textShape.setAnchor(anchor);
    }

    private static void changeFontSize(final XSLFTextShape textShape) {
        textShape.setTextAutofit(TextShape.TextAutofit.SHAPE);
        textShape.getTextParagraphs().forEach(paragraph -> {
            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                textRun.setFontSize(28.);
            }
        });
    }

    private static void memStats() {
        long heapSize = Runtime.getRuntime().totalMemory();
        // Get amount of free memory within the heap in bytes. This size will increase
        // after garbage collection and decrease as new objects are created.
        long heapFreeSize = Runtime.getRuntime().freeMemory();

        log.info("heap size {}; heap Free size {}; used heap {}", formatSize(heapSize), formatSize(heapFreeSize),
                formatSize(heapSize - heapFreeSize));

    }

    private static String formatSize(final long v) {
        if (v < 1024) return v + " B";
        int z = (63 - Long.numberOfLeadingZeros(v)) / 10;
        return String.format("%.1f %sB", (double) v / (1L << (z * 10)), " KMGTPE".charAt(z));
    }

}
