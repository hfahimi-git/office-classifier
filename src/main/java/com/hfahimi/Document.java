package com.hfahimi;

import java.io.IOException;

abstract class Document {

    abstract String getAuthor();
    abstract String getLastModifier();
    abstract void close() throws IOException;

    public static Document get(String file) throws IOException {
        if (file == null) {
            throw new IOException("invalid type");
        }
        if (file.endsWith(".docx") || file.endsWith(".doc")) {
            return new Word(file);
        } else if (file.endsWith(".xlsx") || file.endsWith(".xls")) {
            return new Excel(file);
        } else {
            throw new IOException("invalid type");
        }

    }
}
