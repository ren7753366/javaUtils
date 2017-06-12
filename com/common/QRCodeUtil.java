package com.common;

import java.awt.Image;
import java.awt.geom.AffineTransform;
import java.awt.image.AffineTransformOp;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Hashtable;

import javax.imageio.ImageIO;

import org.apache.commons.lang.StringUtils;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

/**
 * 
 * <p>
 * Title:QRCodeUtil.java
 * </p>
 * <p>
 * Description:二维码生成工具类
 * </p>
 * @author renms
 * @version 1.0
 */
public class QRCodeUtil {
    private static final int BLACK = 0xFF000000;
    private static final int WHITE = 0xFFFFFFFF;

    /**
     * 
     * <p>
     * Title:writeImage
     * </p>
     * <p>
     * Description:生成二维码图片
     * </p>
     * 
     * @param content
     *            二维码图片中显示的信息
     * @param imgType
     *            图片类型
     * @param width
     *            宽度
     * @param height
     *            高度
     * @return 图片字节数组
     * @throws Exception
     *             异常
     */
    public static byte[] writeImage(String content, String imgType, int width,
            int height) throws Exception {
        if ("xtl".equals(imgType)) {
            return writeIconImage(content, width, height, "xtl.PNG");
        } else {
            return writeSimpleImage(content, width, height);
        }
    }

    /**
     * 
     * <p>
     * Title:writeSimpleImage
     * </p>
     * <p>
     * Description: 生成简单二维码，及普通的纯二维码
     * </p>
     * 
     * @param content 二维码蕴含信息
     * @param width 二维码宽度
     * @param height 二维码高度
     * @return 二维码图片字节数组
     * @throws Exception 异常
     */
    public static byte[] writeSimpleImage(String content, int width, int height) throws Exception {
        MultiFormatWriter multiFormatWriter = new MultiFormatWriter();
        Hashtable hints = new Hashtable();
        hints.put(EncodeHintType.MARGIN, 0);
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
        BitMatrix bitMatrix;
        ByteArrayOutputStream bo = new ByteArrayOutputStream();
        bitMatrix = multiFormatWriter.encode(content, BarcodeFormat.QR_CODE,
                width, height, hints);
        MatrixToImageWriter.writeToStream(bitMatrix, "jpg", bo);
        bo.write(new byte[3000]);
        return bo.toByteArray();
    }

   /**
    * 
   * <p>Title:writeIconImage</p>
   * <p>Description:生成中间带有logo的二维码图片</p>
   * @param content 二维码蕴含信息
   * @param width 宽度
   * @param height 高度
   * @param png logo图片的文件路径
   * @return 二维码图片字节数组
   * @throws Exception 异常
    */
    public static byte[] writeIconImage(String content, int width, int height,
            String png) throws Exception {
        ByteArrayOutputStream bo = new ByteArrayOutputStream();
        InputStream srcImgStream = QRCodeUtil.class.getClassLoader()
                .getResourceAsStream("icon/" + png);
        BufferedImage srcImage = ImageIO.read(srcImgStream);
        BufferedImage img = writeIconCodeImage(content, width, height,
                srcImage, 3);
        ImageIO.write(img, "jpg", bo);
        return bo.toByteArray();
    }

    /**
     * 得到BufferedImage
     * 
     * @param content
     *            二维码显示的文本
     * @param width
     *            二维码的宽度
     * @param height
     *            二维码的高度
     * @param srcImg
     *            中间嵌套的图片
     * @param multi 整个二维码大小相对于logo大小的倍数           
     * @return 带有logo的二维码图片
     * @throws WriterException 异常
     * @throws IOException 异常
     */
    private static BufferedImage writeIconCodeImage(String content, int width, int height, 
            BufferedImage srcImg, int multi) throws WriterException, IOException {

        double ratio = (double) Math.min(width, height) / multi
                / Math.max(srcImg.getWidth(), srcImg.getHeight());
        int iconTargetWidth = (int) (srcImg.getWidth() * ratio);
        int iconTargetHeight = (int) (srcImg.getHeight() * ratio);

        Image targeIconImage = srcImg.getScaledInstance(iconTargetWidth,
                iconTargetHeight, BufferedImage.SCALE_SMOOTH);
        AffineTransformOp op = new AffineTransformOp(
                AffineTransform.getScaleInstance(ratio, ratio), null);
        targeIconImage = op.filter(srcImg, null); // 缩放后的图标大小

        // 绘制内容
        Hashtable hints = new Hashtable();
        hints.put(EncodeHintType.MARGIN, 0);
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
        MultiFormatWriter multiFormatWriter = new MultiFormatWriter();
        BitMatrix matrix = multiFormatWriter.encode(content,
                BarcodeFormat.QR_CODE, width, height, hints);
        int resultWidth = matrix.getWidth();
        int resultHeight = matrix.getHeight();

        BufferedImage resultImage = new BufferedImage(width, height,
                BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < width; x++) {
            for (int y = 0; y < height; y++) {
                resultImage.setRGB(x, y, matrix.get(x, y) ? BLACK : WHITE);
            }
        }

        int iconPosX = (resultWidth - iconTargetWidth) / 2;
        int iconPosY = (resultHeight - iconTargetHeight) / 2;

        for (int i = iconPosX; i < (iconPosX + iconTargetWidth); i++) {
            for (int j = iconPosY; j < (iconPosY + iconTargetHeight); j++) {
                resultImage.setRGB(
                        i,
                        j,
                        ((BufferedImage) targeIconImage).getRGB(i - iconPosX, j
                                - iconPosY));
            }
        }
        return resultImage;
    }

}
