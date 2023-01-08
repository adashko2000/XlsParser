package org.insert.parser.sheets;

import java.util.ArrayList;
import java.util.List;

public class TemplateSheet {
    public String filePath;

    public String tableName;

    public String outputFile;

    public List<String> rowIgnoringValues = new ArrayList<>();

    public List<SheetColumn> columns = new ArrayList<>();
}
