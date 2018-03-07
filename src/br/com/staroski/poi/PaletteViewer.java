package br.com.staroski.poi;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Container;
import java.awt.GridLayout;
import java.awt.SystemColor;
import java.lang.reflect.Field;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeSet;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.SwingConstants;
import javax.swing.UIManager;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

/**
 * This is a utility allows visualization of Apache POI's color palette and HSSFColor constants.
 * 
 * @author Ricardo Artur Staroski
 */
public final class PaletteViewer extends JFrame {

    /** Used to sort colors by HSB (Hue Saturation Brightness). */
    private static final class HSBComparator implements Comparator<HSSFColor> {

        @Override
        public int compare(HSSFColor a, HSSFColor b) {
            short[] rgb1 = a.getTriplet();
            short[] rgb2 = b.getTriplet();
            float[] hsb1 = Color.RGBtoHSB(rgb1[0], rgb1[1], rgb1[2], new float[3]);
            float[] hsb2 = Color.RGBtoHSB(rgb2[0], rgb2[1], rgb2[2], new float[3]);
            if (hsb1[0] < hsb2[0]) {
                return -1;
            }
            if (hsb1[0] > hsb2[0]) {
                return 1;
            }
            if (hsb1[1] < hsb2[1]) {
                return -1;
            }
            if (hsb1[1] > hsb2[1]) {
                return 1;
            }
            if (hsb1[2] < hsb2[2]) {
                return -1;
            }
            if (hsb1[2] > hsb2[2]) {
                return 1;
            }
            return 0;
        }
    }

    private static final long serialVersionUID = 1;

    /** Map of constant names */
    private static final Map<Color, List<String>> COLOR_NAMES;

    static {
        try {
            COLOR_NAMES = new HashMap<>();
            Class<?>[] innerClasses = HSSFColor.class.getDeclaredClasses();
            for (Class<?> inner : innerClasses) {
                if (HSSFColor.class.isAssignableFrom(inner)) {
                    Field field = null;
                    short[] triplet = null;
                    try {
                        field = inner.getDeclaredField("triplet");
                        field.setAccessible(true);
                        triplet = (short[]) field.get(null);
                    } catch (NoSuchFieldException e) {
                        field = inner.getDeclaredField("instance");
                        field.setAccessible(true);
                        HSSFColor poiColor = (HSSFColor) field.get(null);
                        triplet = poiColor.getTriplet();
                    }
                    Color color = new Color(triplet[0], triplet[1], triplet[2]);
                    String name = inner.getSimpleName();
                    List<String> list = COLOR_NAMES.get(color);
                    if (list == null) {
                        list = new LinkedList<>();
                        COLOR_NAMES.put(color, list);
                    }
                    list.add(name);
                }
            }

        } catch (Throwable t) {
            throw new ExceptionInInitializerError(t);
        }
    }

    /**
     * Program's main entry point
     * 
     * @param args
     *            not used
     */
    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Throwable t) {
            t.printStackTrace();
        }
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFPalette palette = workbook.getCustomPalette();
            PaletteViewer paletteViewer = new PaletteViewer(palette);
            paletteViewer.setDefaultCloseOperation(EXIT_ON_CLOSE);
            paletteViewer.setLocationRelativeTo(null);
            paletteViewer.setVisible(true);
        } catch (Throwable t) {
            t.printStackTrace();
        }
    }

    /** private constructor */
    private PaletteViewer(HSSFPalette palette) {
        super(HSSFColor.class.getSimpleName());
        setSize(640, 480);
        JPanel panel = new JPanel(new GridLayout(0, 6));
        panel.setOpaque(true);
        panel.setBackground(SystemColor.control);
        Container content = getContentPane();
        content.setLayout(new BorderLayout());
        content.add(panel, BorderLayout.CENTER);

        Set<HSSFColor> colors = getSortedColors(palette);
        addColorsToPanel(colors, panel);
    }

    /** adds the colors to the panel */
    private void addColorsToPanel(Set<HSSFColor> colors, JPanel panel) {
        for (HSSFColor color : colors) {
            short[] triplet = color.getTriplet();
            Color javaColor = new Color(triplet[0], triplet[1], triplet[2]);
            int rgb = javaColor.getRGB() & 0x00FFFFFF;
            String hexa = String.format("%06X", rgb);
            JLabel label = new JLabel(hexa, SwingConstants.CENTER);
            label.setForeground(rgb < 0x808080 ? Color.WHITE : Color.BLACK);
            StringBuilder names = new StringBuilder("<html><p align='center'>");
            int count = 0;
            for (String name : findNames(color)) {
                if (count > 0) {
                    names.append("<br/>or<br/>");
                }
                names.append(name);
                count++;
            }
            names.append("</p></html>");
            label.setToolTipText(names.toString());
            label.setOpaque(true);
            label.setBackground(javaColor);
            panel.add(label);
        }
    }

    /** finds the list of contants that correspond to the specified color */
    private List<String> findNames(HSSFColor poiColor) {
        short[] triplet = poiColor.getTriplet();
        Color color = new Color(triplet[0], triplet[1], triplet[2]);
        for (Entry<Color, List<String>> entry : COLOR_NAMES.entrySet()) {
            if (color.equals(entry.getKey())) {
                return Collections.unmodifiableList(entry.getValue());
            }
        }
        return Collections.emptyList();
    }

    /** gets the palette colors sorted by HSB (Hue Saturation Brightness) */
    private Set<HSSFColor> getSortedColors(HSSFPalette palette) {
        Set<HSSFColor> sortedColors = new TreeSet<>(new HSBComparator());
        short index = 0x08;
        HSSFColor poiColor = palette.getColor(index);
        while (poiColor != null) {
            sortedColors.add(poiColor);
            index++;
            poiColor = palette.getColor(index);
        }
        return sortedColors;
    }
}
