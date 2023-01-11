package org.insert.parser.sheets;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.ArrayList;
import java.util.List;

public class TemplateSheet {
    public String filePath;

    public String tableName;

    public String outputFile;
    @JsonIgnoreProperties(ignoreUnknown = true)
    public List<String> rowIgnoringValues = new ArrayList<>();
    @JsonIgnoreProperties(ignoreUnknown = true)
    public List<String> rowIgnoringHeaders = new ArrayList<>();

    public List<SheetColumn> columns = new ArrayList<>();
}
