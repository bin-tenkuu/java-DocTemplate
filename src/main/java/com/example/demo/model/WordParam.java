package com.example.demo.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.util.Units;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

public sealed interface WordParam {

    static WordParam text(Object msg) {
        return new Text(String.valueOf(msg));
    }

    static WordParam image(InputStream inputStream, int width, int height) {
        return new Image(inputStream, width, height);
    }

    static WordParam image(BufferedImage image) throws IOException {
        return new Image(image);
    }

    static WordParam image(File file) throws IOException {
        return new Image(ImageIO.read(file));
    }

    @Getter
    @AllArgsConstructor
    final class Text implements WordParam {
        private final String msg;
    }

    @Getter
    @AllArgsConstructor
    final class Image implements WordParam {
        private final InputStream inputStream;
        private final int width;
        private final int height;

        public Image(BufferedImage bufferedImage) throws IOException {
            this.inputStream = toInputStream(bufferedImage);
            this.width = Units.toEMU(bufferedImage.getWidth());
            this.height = Units.toEMU(bufferedImage.getHeight());
        }

        private static InputStream toInputStream(BufferedImage image) throws IOException {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            ImageIO.write(image, "jpg", os);
            return new ByteArrayInputStream(os.toByteArray());
        }
    }

}
