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

import java.awt.Dimension;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import javax.swing.ImageIcon;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.border.EmptyBorder;
import net.miginfocom.swing.MigLayout;

/**
 * The app GUI.
 * @author Paulo Roberto Massa Cereda
 */
public class MainFrame {

    private JDialog theDialog;

    public MainFrame() {
        theDialog = new JDialog();
        theDialog.setTitle("excel2datatool");
        theDialog.setMinimumSize(new Dimension(128, 128));
        theDialog.setResizable(false);
        theDialog.setModal(true);

        theDialog.addWindowListener(new WindowAdapter() {

            @Override
            public void windowClosing(WindowEvent e) {
                try {
                    theDialog.dispose();
                } catch (Exception ex) {
                }
            }
        });

        theDialog.setLayout(new MigLayout());
        JLabel dropHere = new JLabel(new ImageIcon(MainFrame.class.getResource("/excel2datatool/imgs/drophere.png")));
        new FileDrop(dropHere, new EmptyBorder(0, 0, 0, 0), new FileDrop.Listener() {

            @Override
            public void filesDropped(File[] files) {
                try {
                    File theFile = null;
                    theFile = files[0];
                    CSVConverter converter = new CSVConverter(theFile);
                    converter.setComponent(theDialog);
                    if (!converter.isValid()) {
                        JOptionPane.showMessageDialog(theDialog, "I'm sorry, the file appears to be invalid.", "Boo!", JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    if (!converter.save()) {
                        JOptionPane.showMessageDialog(theDialog, "Something bad happened. I blame the soup.", "Boo!", JOptionPane.ERROR_MESSAGE);
                    }
                } catch (Exception e) {
                }
            }
        });

        theDialog.add(dropHere, "grow");
        theDialog.pack();
    }

    public void show() {
        theDialog.setLocationRelativeTo(null);
        theDialog.setAlwaysOnTop(true);
        theDialog.setVisible(true);
    }
}
