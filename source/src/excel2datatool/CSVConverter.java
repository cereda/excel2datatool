/**
 * excel2datatool -- Excel spreadsheets to CSV + LaTeX + datatool
 * Copyright (c) 2012, Paulo Roberto Massa Cereda All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice,
 * this list of conditions and the following disclaimer.
 *
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 * this list of conditions and the following disclaimer in the documentation
 * and/or other materials provided with the distribution.
 *
 * 3. Neither the name of the project's author nor the names of its contributors
 * may be used to endorse or promote products derived from this software without
 * specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 *
 */
package excel2datatool;

import au.com.bytecode.opencsv.CSVWriter;
import java.awt.Component;
import java.io.*;
import java.util.ArrayList;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

/**
 * The CSV converter.
 * 
 * @author Paulo Roberto Massa Cereda
 */
public class CSVConverter {

    private File theFile;
    private Component theComponent;

    public void setComponent(Component theComponent) {
        this.theComponent = theComponent;
    }

    public CSVConverter(File theFile) {
        this.theFile = theFile;
        theComponent = null;
    }

    public boolean isValid() {
        if (!theFile.isFile()) {
            return false;
        }
        return isExcel(theFile.getName());
    }

    private boolean isExcel(String fileName) {
        if ((fileName.toLowerCase().endsWith(".xls"))
                || (fileName.toLowerCase().endsWith(".xlsx"))) {
            return true;
        }
        return false;
    }

    public boolean save() {
        try {
            InputStream input = new FileInputStream(theFile);
            Workbook workBook = WorkbookFactory.create(input);
            Sheet theSheet = workBook.getSheetAt(0);
            ArrayList<String> theRow = new ArrayList<String>();

            File theCSVFile = Utils.getFileToSave("Comma separated value", "csv", theComponent);
            if (theCSVFile == null) {
                System.out.println("No valid CSV file.");
                return false;
            }

            File theLaTeXFile = null;
            if (JOptionPane.showConfirmDialog(theComponent, "Do you want to generate the LaTeX file too?", "Woah cowboy.", JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE) == JOptionPane.YES_OPTION) {
                theLaTeXFile = Utils.getFileToSave("LaTeX file", "tex", theComponent);
                if (theLaTeXFile == null) {
                    System.out.println("No valid LaTeX file.");
                    return false;
                }
            }

            CSVWriter writerCSV = new CSVWriter(new FileWriter(theCSVFile));
            for (Row theCurrentRow : theSheet) {
                theRow.clear();
                for (Cell theCell : theCurrentRow) {
                    theRow.add(theCell.toString());
                }
                writerCSV.writeNext((String[]) theRow.toArray(new String[theRow.size()]));
            }
            writerCSV.close();

            JOptionPane.showMessageDialog(theComponent, "The CSV file was generated successfully!", "Done!", JOptionPane.INFORMATION_MESSAGE);

            if (theLaTeXFile != null) {
                LaTeXWriter writerLaTeX = new LaTeXWriter(theLaTeXFile, theCSVFile);
                if (writerLaTeX.generate()) {
                    JOptionPane.showMessageDialog(theComponent, "The LaTeX file was generated successfully!", "Done!", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    return false;
                }
            }
        } catch (IOException ioe) {
            System.out.println("Error: " + ioe.getMessage());
            return false;
        } catch (InvalidFormatException ife) {
            System.out.println("Error: " + ife.getMessage());
            return false;
        }
        return true;
    }
}
