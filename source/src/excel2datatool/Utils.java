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

import java.awt.Component;
import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 * Provides helper methods.
 *
 * @author Paulo Roberto Massa Cereda
 */
public class Utils {

    public static File getFileToSave(String description, String extension, Component component) {
        JFileChooser fileChooser = new JFileChooser();
        FileFilter filter = new FileNameExtensionFilter(description, new String[]{extension});
        fileChooser.setFileFilter(filter);
        int action = fileChooser.showSaveDialog(component);
        if (action == JFileChooser.APPROVE_OPTION) {
            if (fileChooser.getSelectedFile() != null) {
                File theFile = fileChooser.getSelectedFile();
                if (!theFile.getAbsolutePath().toLowerCase().endsWith("." + extension)) {
                    theFile = new File(theFile.getAbsolutePath() + "." + extension);
                }
                return theFile;
            }
        }
        return null;
    }
}
