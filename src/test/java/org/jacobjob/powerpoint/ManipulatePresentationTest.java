package org.jacobjob.powerpoint;

import org.junit.jupiter.api.Test;

import java.awt.geom.Rectangle2D;

import static org.jacobjob.powerpoint.ManipulatePresentation.newXPosition;
import static org.jacobjob.powerpoint.ManipulatePresentation.newYPosition;
import static org.junit.jupiter.api.Assertions.assertEquals;

class ManipulatePresentationTest {

    private ManipulatePresentation manipulator = new ManipulatePresentation();

    @Test
    void changePictureAnchorSizeSquare() {
        Rectangle2D anchor = new Rectangle2D.Double();
        anchor.setFrame(0, 0, 100, 100);
        anchor = manipulator.changePictureAnchorSize(anchor);
        assertEquals(Math.floor(540 - newYPosition - 5), anchor.getHeight());
        assertEquals(520, anchor.getWidth());
        assertEquals(anchor.getHeight(), anchor.getWidth());
    }

    @Test
    void changePictureAnchorSizeLonger() {
        Rectangle2D anchor = new Rectangle2D.Double();
        anchor.setFrame(0, 0, 200, 100);
        anchor = manipulator.changePictureAnchorSize(anchor);
        assertEquals(Math.floor(960 - newXPosition - 5), anchor.getWidth());
        assertEquals(Math.floor(anchor.getWidth() / 2), anchor.getHeight());
    }

    @Test
    void changePictureAnchorSizeWider() {
        Rectangle2D anchor = new Rectangle2D.Double();
        anchor.setFrame(0, 0, 100, 200);
        anchor = manipulator.changePictureAnchorSize(anchor);
        assertEquals(Math.floor(540 - newYPosition - 5), anchor.getHeight());
        assertEquals(Math.floor(anchor.getHeight() / 2), anchor.getWidth());
    }
}