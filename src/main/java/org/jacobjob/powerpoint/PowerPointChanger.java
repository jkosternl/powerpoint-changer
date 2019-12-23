package org.jacobjob.powerpoint;

import lombok.extern.log4j.Log4j2;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.time.StopWatch;

import java.io.File;
import java.util.LinkedList;
import java.util.List;

@Log4j2
public class PowerPointChanger {

    private static final String POWERPOINT_FILE = "Opw 42.pptx";
    private static final String POWERPOINT_FILE2 = "Opw 585.pptx";

//    private static final Logger log = LogManager.getLogger(PowerPointChanger.class.getName());

    public static void main(String[] args) {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        log.info("Starting");
        memStats();

        List<File> fileList = getFileListFromPath(new File("D:/Temp/Databank"));
        log.info("Found {}", fileList.size());

        ManipulatePresentation manipulator = new ManipulatePresentation();
        int failedCount = 0, goodCount = 0;
        for (File file : fileList) {
            try {
                log.info("Stats: Good: {}. Failed: {}", goodCount, failedCount);
                manipulator.processPowerpointFile(file);
                goodCount++;
            } catch (Exception e){
                log.error("Failed processing: {}", file.getName());
                failedCount++;
            }
        }
//        manipulator.processPowerpointFile(new File(POWERPOINT_FILE));

        memStats();
        stopWatch.stop();
        log.info("Programma beeindigd na: {}", stopWatch);
    }

    private static List<File> getFileListFromPath(final File directory) {
        return (LinkedList<File>) FileUtils.listFiles(directory, new String[]{"pptx"}, true);
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
