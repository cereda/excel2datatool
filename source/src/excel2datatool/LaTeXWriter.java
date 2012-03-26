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

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

/**
 * The LaTeX writer.
 * 
 * @author Paulo Roberto Massa Cereda
 */
public class LaTeXWriter {

    private File theLaTeXFile;
    private File theCSVFile;

    public LaTeXWriter(File theFile, File theCSVFile) {
        this.theLaTeXFile = theFile;
        this.theCSVFile = theCSVFile;
    }

    public boolean generate() {
        try {
            FileWriter writer = new FileWriter(theLaTeXFile);
            writer.write("\\documentclass{article}\n\n");
            writer.write("\\usepackage[T1]{fontenc}\n");
            writer.write("\\usepackage[utf8]{inputenc}\n\n");
            writer.write("\\usepackage{datatool}\n\n");
            writer.write("\\DTLloaddb{mydb}{" + theCSVFile.getName() + "}\n\n");
            writer.write("\\begin{document}\n\n");
            writer.write("\\DTLdisplaydb{mydb}\n\n");
            writer.write("\\end{document}");
            writer.close();
        } catch (IOException ioe) {
            return false;
        }
        return true;
    }
}
