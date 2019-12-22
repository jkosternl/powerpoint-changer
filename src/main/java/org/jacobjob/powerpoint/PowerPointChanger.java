package org.jacobjob.powerpoint;

import org.apache.commons.lang3.time.StopWatch;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class PowerPointChanger {

    private static final String POWERPOINT_FILE = "Opw 42.pptx";
    private static final String POWERPOINT_FILE2 = "Opw 585.pptx";

    private static final Logger log = LogManager.getLogger(PowerPointChanger.class.getName());

    public static void main(String[] args) throws Exception {
        StopWatch stopWatch = new StopWatch();
        stopWatch.start();
        log.info("Starting");
        memStats();

        ManipulatePresentation manipulator = new ManipulatePresentation();
        manipulator.processPowerpointFile(POWERPOINT_FILE);

        memStats();
        stopWatch.stop();
        log.info("Programma beeindigd na: {}", stopWatch);
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
