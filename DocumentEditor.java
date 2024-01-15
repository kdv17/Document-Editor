package just;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Scanner;
import java.util.List;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.event.*;
import javax.swing.filechooser.*;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Document;
import javax.swing.text.Element;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;

import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.graphics.PDXObject;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;


public class DocumentEditor extends JFrame implements ActionListener {
    JTextPane textPane;
    JScrollPane scrollPane;
    JLabel fontLabel;
    JSpinner fontSizeSpinner;
    JButton fontColor;
    JButton boldButton;
    JButton italicButton;
    JButton underlineButton;
    JButton alignLeftButton;
    JButton alignCenterButton;
    JButton alignRighButton;
    JButton bulletButton;
    JComboBox fontBox;
    JMenuBar menuBar;
    JMenu fileMenu;
    JMenuItem openItem;
    JMenuItem saveItem;
    JMenuItem closeItem;
    JToolBar toolBar;

    DocumentEditor() {
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setTitle("Document Editor");
        //setSize(500, 500);
        setExtendedState(MAXIMIZED_BOTH);
        setLocationRelativeTo(null);
        
        setLayout(new BorderLayout());

        // Set the color scheme
        Color bgColor = new Color(0, 255, 255);
        Color fgColor = new Color(50, 50, 50);
        getContentPane().setBackground(bgColor);
        setBackground(bgColor);
        textPane = new JTextPane();

        
        textPane.setFont(new Font("Arial", Font.PLAIN, 20));
        textPane.setBackground(bgColor);
        textPane.setForeground(fgColor);

        scrollPane = new JScrollPane(textPane);
        scrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);

        fontLabel = new JLabel("Font: ");
        //fontLabel.setOpaque(true);    

        fontSizeSpinner = new JSpinner();
        fontSizeSpinner.setPreferredSize(new Dimension(50, 25));
        fontSizeSpinner.setValue(20);
        fontSizeSpinner.addChangeListener(new ChangeListener() {
            public void stateChanged(ChangeEvent e) {
                textPane.setFont(new Font(textPane.getFont().getFamily(), Font.PLAIN, (int) fontSizeSpinner.getValue()));
            }
        });

        fontColor = new JButton("Color");
        fontColor.setBackground(Color.YELLOW);
        fontColor.setForeground(Color.BLACK);
        fontColor.addActionListener(this);

        boldButton = new JButton("BOLD");
        boldButton.setBackground(Color.BLACK);
        boldButton.setForeground(Color.WHITE);
        boldButton.setBorderPainted(false);
        boldButton.addActionListener(this);

        italicButton = new JButton("Italic");
        italicButton.setBackground(Color.BLACK);
        italicButton.setForeground(Color.WHITE);
        italicButton.setBorderPainted(false);
        italicButton.addActionListener(this);

        underlineButton = new JButton("Underline");
        underlineButton.setBackground(Color.BLACK);
        underlineButton.setForeground(Color.WHITE);
        underlineButton.setBorderPainted(false);
        underlineButton.addActionListener(this);

        alignLeftButton = new JButton("AlignL");
        alignLeftButton.setBackground(Color.BLACK);
        alignLeftButton.setForeground(Color.WHITE);
        alignLeftButton.setBorderPainted(false);
        alignLeftButton.addActionListener(this);

        alignCenterButton = new JButton("AlignC");
        alignCenterButton.setBackground(Color.BLACK);
        alignCenterButton.setForeground(Color.WHITE);
        alignCenterButton.setBorderPainted(false);
        alignCenterButton.addActionListener(this);

        alignRighButton = new JButton("AlignR");
        alignRighButton.setBackground(Color.BLACK);
        alignRighButton.setForeground(Color.WHITE);
        alignRighButton.setBorderPainted(false);
        alignRighButton.addActionListener(this);

        bulletButton = new JButton("Bullets");
        bulletButton.setBackground(Color.BLACK);
        bulletButton.setForeground(Color.WHITE);
        bulletButton.setBorderPainted(false);
        bulletButton.addActionListener(this);





        String[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();

        fontBox = new JComboBox(fonts);
        fontBox.addActionListener(this);
        fontBox.setSelectedItem("Arial");

        menuBar = new JMenuBar();
        fileMenu = new JMenu("File");
        openItem = new JMenuItem("Open");
        saveItem = new JMenuItem("Save");
        closeItem = new JMenuItem("Close");

        openItem.addActionListener(this);
        saveItem.addActionListener(this);
        closeItem.addActionListener(this);

        fileMenu.add(openItem);
        fileMenu.add(saveItem);
        fileMenu.add(closeItem);
        menuBar.add(fileMenu);

        toolBar = new JToolBar();
        toolBar.add(boldButton);
        toolBar.add(italicButton);
        toolBar.add(underlineButton);
        toolBar.add(alignLeftButton);
        toolBar.add(alignCenterButton);
        toolBar.add(alignRighButton);
        toolBar.add(bulletButton);
        add(toolBar, BorderLayout.NORTH);


        setJMenuBar(menuBar);
        // Add components to the JFrame using the BorderLayout
        add(scrollPane, BorderLayout.CENTER);

        JPanel topPanel = new JPanel();
        topPanel.add(fontLabel);
        topPanel.add(fontSizeSpinner);
        topPanel.add(fontColor);
        topPanel.add(fontBox);
        
        JPanel northPanel = new JPanel(new FlowLayout());
        northPanel.add(toolBar, BorderLayout.NORTH);
        northPanel.add(topPanel, BorderLayout.SOUTH);

        add(northPanel, BorderLayout.NORTH);
        northPanel.setBackground(fgColor);



        setVisible(true);
    }

    public void actionPerformed(ActionEvent e) {
      if (e.getSource() == fontColor) {
          JColorChooser colorChooser = new JColorChooser();
          Color color = JColorChooser.showDialog(null, "Choose a color", Color.black);
          textPane.setForeground(color);
      }
      if (e.getSource() == fontBox) {
          textPane.setFont(new Font((String) fontBox.getSelectedItem(), Font.PLAIN, textPane.getFont().getSize()));
      }
      if(e.getSource()==boldButton){
          int start = textPane.getSelectionStart();
          int end = textPane.getSelectionEnd();
          if (start == end) { // No selection, cursor position.
              return;
          }
          StyledDocument doc = textPane.getStyledDocument();
          AttributeSet oldAttr = doc.getCharacterElement(start).getAttributes();
          SimpleAttributeSet newAttr = new SimpleAttributeSet(oldAttr);
          StyleConstants.setBold(newAttr, !StyleConstants.isBold(oldAttr));
          doc.setCharacterAttributes(start, end - start, newAttr, true);
      }
      if(e.getSource()==italicButton){
          int start = textPane.getSelectionStart();
          int end = textPane.getSelectionEnd();
          if (start == end) { // No selection, cursor position.
              return;
          }
          StyledDocument doc = textPane.getStyledDocument();
          AttributeSet oldAttr = doc.getCharacterElement(start).getAttributes();
          SimpleAttributeSet newAttr = new SimpleAttributeSet(oldAttr);
          StyleConstants.setItalic(newAttr, !StyleConstants.isItalic(oldAttr));
          doc.setCharacterAttributes(start, end - start, newAttr, true);
      }
      if(e.getSource()==underlineButton){
        int start = textPane.getSelectionStart();
        int end = textPane.getSelectionEnd();
        if (start == end) { // No selection, cursor position.
            return;
        }
        StyledDocument doc = textPane.getStyledDocument();
        AttributeSet oldAttr = doc.getCharacterElement(start).getAttributes();
        SimpleAttributeSet newAttr = new SimpleAttributeSet(oldAttr);
        StyleConstants.setUnderline(newAttr, !StyleConstants.isUnderline(oldAttr));
        doc.setCharacterAttributes(start, end - start, newAttr, true);
      }
    if(e.getSource()==alignLeftButton){
        // Code to align the text left
        StyledDocument doc = textPane.getStyledDocument();
        SimpleAttributeSet newAttr = new SimpleAttributeSet();
        StyleConstants.setAlignment(newAttr, StyleConstants.ALIGN_LEFT);
        doc.setParagraphAttributes(0, doc.getLength(), newAttr, false);
      }
      if(e.getSource()==alignCenterButton){
         // Code to align the text center
         StyledDocument doc = textPane.getStyledDocument();
         SimpleAttributeSet newAttr = new SimpleAttributeSet();
         StyleConstants.setAlignment(newAttr, StyleConstants.ALIGN_CENTER);
         doc.setParagraphAttributes(0, doc.getLength(), newAttr, false);
      }
      if(e.getSource()==alignRighButton){
        // Code to align the text right
        StyledDocument doc = textPane.getStyledDocument();
        SimpleAttributeSet newAttr = new SimpleAttributeSet();
        StyleConstants.setAlignment(newAttr, StyleConstants.ALIGN_RIGHT);
        doc.setParagraphAttributes(0, doc.getLength(), newAttr, false);
      }
      if(e.getSource()==bulletButton){
        int start = textPane.getSelectionStart();
        int end = textPane.getSelectionEnd();
        if (start == end) { // No selection, cursor position.
            return;
        }
        try {
            StyledDocument doc = textPane.getStyledDocument();
            String selectedText = doc.getText(start, end - start);
            String[] lines = selectedText.split("\n");
            for (int i = lines.length - 1; i >= 0; i--) {
                String lineStart = lines[i].split("\\S", 2)[0];
                int lineStartPos = start + selectedText.indexOf(lines[i]);
                int lineEndPos = lineStartPos + lines[i].length();
                if (lineStart.contains("•")) {
                    // Remove bullet point
                    doc.remove(lineStartPos + lineStart.indexOf("•"), 3);
                } else {
                    // Add bullet point
                    doc.insertString(lineStartPos + lineStart.length(), "»", null);
                }
            }
            textPane.select(start, end);
        } catch (BadLocationException ex) {
            ex.printStackTrace();
        }
      }
      if (e.getSource() == openItem) {
          JFileChooser fileChooser = new JFileChooser();
          fileChooser.setCurrentDirectory(new File("."));
          FileNameExtensionFilter filter = new FileNameExtensionFilter("Text, PDF and WORD files", "txt", "pdf", "docx");
          fileChooser.setFileFilter(filter);

          int response = fileChooser.showOpenDialog(null);
          if (response == JFileChooser.APPROVE_OPTION) {
              File file = new File(fileChooser.getSelectedFile().getAbsolutePath());
              String fileName = file.getName();
              String extension = fileName.substring(fileName.lastIndexOf(".") + 1);

              if (extension.equals("txt")) {
                  Scanner fileIn = null;
                  try {
                      fileIn = new Scanner(file);
                      if (file.isFile()) {
                          while (fileIn.hasNextLine()) {
                              String line = fileIn.nextLine() + "\n";
                              textPane.setText(textPane.getText() + line);
                          }
                      }
                  } catch (FileNotFoundException e1) {
                      e1.printStackTrace();
                  } finally {
                      fileIn.close();
                  }
              } else if (extension.equals("pdf")) {
                PDDocument document = null;
                try {
                    document = PDDocument.load(file);
                    PDFTextStripper stripper = new PDFTextStripper();
                    String text = stripper.getText(document);
                    textPane.setText(text);
                    PDPageTree pages = document.getDocumentCatalog().getPages();
                    for (PDPage page : pages) {
                        for (COSName name : page.getResources().getXObjectNames()) {
                            PDXObject xobject = page.getResources().getXObject(name);
                            if (xobject instanceof PDImageXObject) {
                                PDImageXObject image = (PDImageXObject) xobject;
                                BufferedImage bufferedImage = image.getImage();
                                ImageIcon icon = new ImageIcon(bufferedImage);
                                textPane.setCaretPosition(textPane.getDocument().getLength());
                                textPane.insertIcon(icon);
                                textPane.setCaretPosition(0);
                            }
                        }
                    }
            
                    // Add the MouseListener to the JTextPane here
                    textPane.addMouseListener(new MouseAdapter() {
                        public void mouseClicked(MouseEvent e) {
                            // Check if the user clicked on an image
                            // If so, display a pop-up menu with options to resize, delete or insert a new image
                            int pos = textPane.viewToModel2D(e.getPoint());
                            StyledDocument doc = (StyledDocument) textPane.getDocument();
                            Element imageElement = doc.getCharacterElement(pos);
                            // Check if the user clicked on an image
                            if (StyleConstants.getIcon(imageElement.getAttributes()) != null) {
                                // If so, display a pop-up menu with options to resize, delete or insert a new image
                                JPopupMenu popupMenu = new JPopupMenu();
                                JMenuItem resizeItem = new JMenuItem("Resize");
                                resizeItem.addActionListener(new ActionListener() {
                                    public void actionPerformed(ActionEvent e) {
                                        ImageIcon icon = (ImageIcon) StyleConstants.getIcon(imageElement.getAttributes());
                                        Image image = icon.getImage();
                                        // Display a dialog that allows the user to enter the new dimensions for the image
                                        JTextField widthField = new JTextField(5);
                                        JTextField heightField = new JTextField(5);
                                        JPanel myPanel = new JPanel();
                                        myPanel.add(new JLabel("Width:"));
                                        myPanel.add(widthField);
                                        myPanel.add(Box.createHorizontalStrut(15)); // a spacer
                                        myPanel.add(new JLabel("Height:"));
                                        myPanel.add(heightField);
                                        int result = JOptionPane.showConfirmDialog(null, myPanel,
                                                "Enter new dimensions", JOptionPane.OK_CANCEL_OPTION);
                                        if (result == JOptionPane.OK_OPTION) {
                                            // Get the new dimensions from the text fields
                                            int newWidth = Integer.parseInt(widthField.getText());
                                            int newHeight = Integer.parseInt(heightField.getText());
                                            // Use the Image class to resize the image
                                            Image newImage = image.getScaledInstance(newWidth, newHeight, Image.SCALE_SMOOTH);
                                            // Update the image in the JTextPane
                                            // Remove the existing image element
                                            try {
                                                doc.remove(imageElement.getStartOffset(), imageElement.getEndOffset() - imageElement.getStartOffset());
                                            } catch (BadLocationException ex) {
                                                ex.printStackTrace();
                                            }
                                           // Insert the new image at the same position
                                            textPane.setCaretPosition(imageElement.getStartOffset());
                                            textPane.insertIcon(new ImageIcon(newImage));
                                        }
                                    }
                                });
                            popupMenu.add(resizeItem);
            
                            JMenuItem deleteItem = new JMenuItem("Delete");
                            deleteItem.addActionListener(new ActionListener() {
                                @Override
                                public void actionPerformed(ActionEvent e) {
                                    // Delete the image from the JTextPane
                                    // You can use the remove method of the Document class to remove the image element
                                    Document doc = textPane.getDocument();
                                    try {
                                        doc.remove(imageElement.getStartOffset(), imageElement.getEndOffset() - imageElement.getStartOffset());
                                    } catch (BadLocationException ex) {
                                        ex.printStackTrace();
                                    }
                                }
                            });
                            popupMenu.add(deleteItem);
            
                            JMenuItem insertItem = new JMenuItem("Insert");
                            insertItem.addActionListener(new ActionListener() {
                                @Override
                                public void actionPerformed(ActionEvent e) {
                                    // Display a file chooser that allows the user to select an image file
                                    JFileChooser fileChooser = new JFileChooser();
                                    fileChooser.setCurrentDirectory(new File("."));
                                    FileNameExtensionFilter filter = new FileNameExtensionFilter("Image files", "jpg", "png", "gif", "bmp");
                                    fileChooser.setFileFilter(filter);
                                    int response = fileChooser.showOpenDialog(null);
                                    if (response == JFileChooser.APPROVE_OPTION) {
                                        File file = fileChooser.getSelectedFile();
                                        try {
                                            BufferedImage bufferedImage = ImageIO.read(file);
                                            ImageIcon icon = new ImageIcon(bufferedImage);
                                            textPane.insertIcon(icon);
                                        } catch (IOException ex) {
                                            ex.printStackTrace();
                                        }
                                    }
                                }
                            });
                            popupMenu.add(insertItem);
            
                            popupMenu.show(textPane, e.getX(), e.getY());
                        }
                }});
                } catch (IOException ex) {
                    ex.printStackTrace();
                } 
            }
        



            else if (extension.equals("docx")) {
                try {
                    FileInputStream fis = new FileInputStream(file);
                    XWPFDocument document = new XWPFDocument(fis);
                    XWPFWordExtractor extractor = new XWPFWordExtractor(document);
                    String text = extractor.getText();
                    textPane.setText(text);
                    fis.close();
                    // Extract images from the document
                    List<XWPFPictureData> pictures = document.getAllPictures();
                    for (XWPFPictureData picture : pictures) {
                        byte[] data = picture.getData();
                        // Create an image from the data and display it in a label
                        ImageIcon icon = new ImageIcon(data);
                        textPane.insertIcon(icon);
                       
                    }
                    textPane.setCaretPosition(0);

                    textPane.addMouseListener(new MouseAdapter() {
                        public void mouseClicked(MouseEvent e) {
                            // Check if the user clicked on an image
                            // If so, display a pop-up menu with options to resize, delete or insert a new image
                            int pos = textPane.viewToModel2D(e.getPoint());
                            StyledDocument doc = (StyledDocument) textPane.getDocument();
                            Element imageElement = doc.getCharacterElement(pos);
                            // Check if the user clicked on an image
                            if (StyleConstants.getIcon(imageElement.getAttributes()) != null) {
                                // If so, display a pop-up menu with options to resize, delete or insert a new image
                                JPopupMenu popupMenu = new JPopupMenu();
                                JMenuItem resizeItem = new JMenuItem("Resize");
                                resizeItem.addActionListener(new ActionListener() {
                                    public void actionPerformed(ActionEvent e) {
                                        ImageIcon icon = (ImageIcon) StyleConstants.getIcon(imageElement.getAttributes());
                                        Image image = icon.getImage();
                                        // Display a dialog that allows the user to enter the new dimensions for the image
                                        // You can use a JOptionPane to create the dialog
                                        JTextField widthField = new JTextField(5);
                                        JTextField heightField = new JTextField(5);
                                        JPanel myPanel = new JPanel();
                                        myPanel.add(new JLabel("Width:"));
                                        myPanel.add(widthField);
                                        myPanel.add(Box.createHorizontalStrut(15)); // a spacer
                                        myPanel.add(new JLabel("Height:"));
                                        myPanel.add(heightField);
                                        int result = JOptionPane.showConfirmDialog(null, myPanel,
                                                "Enter new dimensions", JOptionPane.OK_CANCEL_OPTION);
                                        if (result == JOptionPane.OK_OPTION) {
                                            // Get the new dimensions from the text fields
                                            int newWidth = Integer.parseInt(widthField.getText());
                                            int newHeight = Integer.parseInt(heightField.getText());
                                            // Use the Image class to resize the image
                                            Image newImage = image.getScaledInstance(newWidth, newHeight, Image.SCALE_SMOOTH);
                                            // Update the image in the JTextPane
                                            // Remove the existing image element
                                            try {
                                                doc.remove(imageElement.getStartOffset(), imageElement.getEndOffset() - imageElement.getStartOffset());
                                            } catch (BadLocationException ex) {
                                                ex.printStackTrace();
                                            }
                                           // Insert the new image at the same position
                                            textPane.setCaretPosition(imageElement.getStartOffset());
                                            textPane.insertIcon(new ImageIcon(newImage));
                                        }
                                    }
                                });
                            popupMenu.add(resizeItem);
                                
                            
                            JMenuItem deleteItem = new JMenuItem("Delete");
                                deleteItem.addActionListener(new ActionListener() {
                                    @Override
                                    public void actionPerformed(ActionEvent e) {
                                        // Delete the image from the JTextPane
                                        // You can use the remove method of the Document class to remove the image element
                                        Document doc = textPane.getDocument();
                                        try {
                                            doc.remove(imageElement.getStartOffset(), imageElement.getEndOffset() - imageElement.getStartOffset());
                                        } catch (BadLocationException ex) {
                                            ex.printStackTrace();
                                        }
                                    }
                                });
                                popupMenu.add(deleteItem);

                                JMenuItem insertItem = new JMenuItem("Insert");
                                insertItem.addActionListener(new ActionListener() {
                                    @Override
                                    public void actionPerformed(ActionEvent e) {
                                        // Display a file chooser that allows the user to select an image file
                                        JFileChooser fileChooser = new JFileChooser();
                                        fileChooser.setCurrentDirectory(new File("."));
                                        FileNameExtensionFilter filter = new FileNameExtensionFilter("Image files", "jpg", "png", "gif", "bmp");
                                        fileChooser.setFileFilter(filter);
                                        int response = fileChooser.showOpenDialog(null);
                                        if (response == JFileChooser.APPROVE_OPTION) {
                                            File file = fileChooser.getSelectedFile();
                                            try {
                                                BufferedImage bufferedImage = ImageIO.read(file);
                                                ImageIcon icon = new ImageIcon(bufferedImage);
                                                textPane.insertIcon(icon);
                                            } catch (IOException ex) {
                                                ex.printStackTrace();
                                            }
                                        }
                                    }
                                });
                                popupMenu.add(insertItem);
                                popupMenu.show(textPane, e.getX(), e.getY());
                            }
                        }
                        });
                    }catch (IOException ex) {
                            ex.printStackTrace();
                        }   
            
                }
            }
        }
                

                /*else if (extension.equals("pptx")) {
                    try {
                        PresentationMLPackage presentationMLPackage = PresentationMLPackage.load(file);
                        StringBuilder sb = new StringBuilder();
                        for (SlidePart slidePart : presentationMLPackage.getMainPresentationPart().getSlideParts()) {
                            List<Object> texts = slidePart.getJAXBNodesViaXPath("//a:t", false);
                            for (Object text : texts) {
                                Text textElement = (Text) ((JAXB)text).getValue();
                                sb.append(textElement.getValue()).append("\n");
                            }
                        }
                        textPane.setText(sb.toString());
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }

                else if (extension.equals("xlsx")) {
                    try {
                        SpreadsheetMLPackage spreadsheetMLPackage = SpreadsheetMLPackage.load(file);
                        // make changes to the document
                        spreadsheetMLPackage.save(new java.io.File("output.xlsx"));
                    } catch (Exception ex) {
                        ex.printStackTrace();
                    }
                }*/
    
        
    
        
            
    

    if(e.getSource()==saveItem) {
        JFileChooser fileChooser1 = new JFileChooser();
        fileChooser1.setCurrentDirectory(new File("."));
    
        int response1 = fileChooser1.showSaveDialog(null);
    
        if(response1 == JFileChooser.APPROVE_OPTION) {
            File file;
            PrintWriter fileOut = null;
        
        file = new File(fileChooser1.getSelectedFile().getAbsolutePath());
        try {
        fileOut = new PrintWriter(file);
        fileOut.println(textPane.getText());
        } 
        catch (FileNotFoundException e1) {
        // TODO Auto-generated catch block
        e1.printStackTrace();
        }
        finally {
        fileOut.close();
        }   
    }
    }
    if(e.getSource()==closeItem) {
    System.exit(0);
    }  
 }
}
