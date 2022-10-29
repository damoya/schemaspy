/*
 * Copyright (C) 2017 Nils Petzaell
 *
 * This file is part of SchemaSpy.
 *
 * SchemaSpy is free software: you can redistribute it and/or modify it under the terms of the GNU
 * Lesser General Public License as published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * SchemaSpy is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
 * even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License along with SchemaSpy. If
 * not, see <http://www.gnu.org/licenses/>.
 */
package org.schemaspy.output.docx;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;
import java.io.File;
import java.io.InputStream;
import java.lang.invoke.MethodHandles;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import org.apache.commons.io.IOUtils;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.Style.BasedOn;
import org.schemaspy.model.Database;
import org.schemaspy.model.ForeignKeyConstraint;
import org.schemaspy.model.Routine;
import org.schemaspy.model.RoutineParameter;
import org.schemaspy.model.Table;
import org.schemaspy.model.TableColumn;
import org.schemaspy.model.TableIndex;
import org.schemaspy.model.View;
import org.schemaspy.output.OutputProducer;

/**
 * @author Nils Petzaell
 */
public class DocxProducer implements OutputProducer {

    private static final Logger LOGGER =
            LoggerFactory.getLogger(MethodHandles.lookup().lookupClass());

    private static final String PK_IMAGE_PATH = "/layout/images/primaryKey.png";

    private static final String FK_IMAGE_PATH = "/layout/images/foreignKey.png";

    private static final String [] TABLE_HEADERS = {
        "Column", "Description"
    };

    private static final String [] TABLE_COLUMN_HEADERS = {
        "Column", "Type", "Size", "Nullable", "Auto", "Default", "Comments"
    };

    private static final String [] VIEW_COLUMN_HEADERS = {
        "Column", "Type", "Size", "Nullable", "Comments"
    };

    private static final String [] FOREIGN_KEY_HEADERS = {
        "Constraint Name", "Child Column", "Parent Column", "Delete Rule"
    };

    private static final String [] INDEXES_HEADERS = {
        "Index Name", "Type Column", "Columns"
    };

    private static final String [] CHECK_HEADERS = {
        "Constraint Name", "Constraint", 
    };

    private static final String [] ROUTINES_HEADERS = {
        "Name", "Type", "Language", "Deterministic", "Return Type", "Security", "Restriction", "Comments"
    };

    private static final String [] ROUTINE_PARAM_HEADERS = {
        "Name", "Type", "Mode"
    };

    private Inline inlinePrimaryKeyImage;

    private Inline inlineForeignKeyImage;

    private ObjectFactory factory = Context.getWmlObjectFactory();

    private WordprocessingMLPackage wordPackage;
    private MainDocumentPart mainDocumentPart;


    // https://www.docx4java.org/docx4j/Docx4j_GettingStarted.pdf

    @Override
    public void generate(Database database, File outputDir) {

        if (database.getTables().isEmpty()) {
            LOGGER.info("No tables to output, nothing written to disk");
            return;
        }

        try {
            createPackage();

            mainDocumentPart.addStyledParagraphOfText("Title", "Database: " + database.getName());

            mainDocumentPart.addStyledParagraphOfText("Heading1", "Schema: " + database.getSchema());

            processTables(database.getTables());

            processViews(database.getViews());

            processRoutines(database.getRoutines());

            File exportFile = new File(outputDir.getPath() + File.separator + "document.docx");

            wordPackage.save(exportFile);
        } catch (Docx4JException e) {
            LOGGER.error("Failed to produce output", e);
        }
    }

    private void processTables(Collection<Table> tables) {

        addTablesList("Tables", tables);
        
        addBr();

        for (Table table : tables) {
            addTable(table);

            addRelationships(table);

            addChecks(table);

            addIndexes(table);

            addBr();
        }
    }


    private void processViews(Collection<View> views) {

        addTablesList("Views", views);

        addBr();

        for (View view : views) {
            addView(view);

            addBr();
        }
    }

    private void processRoutines(Collection<Routine> routines) {

        addRoutinesList(routines);

        addBr();

        for (Routine routine : routines) {
            addRoutine(routine);

            addBr();
        }
    }

    /*
     *
     *  
     * 
     */

    private void addTablesList(String title, Collection<? extends Table> tables) {

        if (CollectionUtils.isEmpty(tables)) {
            return;
        }

        mainDocumentPart.addStyledParagraphOfText("Heading2", title);

        Tbl tbl = createTable(TABLE_HEADERS, tables.size());

        List<Object> rows = tbl.getContent();

        int rowIdx = 0;

        for (Table table : tables) {
            Tr tr = (Tr) rows.get(++rowIdx);

            List<Object> cells = tr.getContent();

            setCellContent(cells, 0, table.getName());
            setCellContent(cells, 1, table.getComments());
        }
    }

    private void addTable(Table table) {

        mainDocumentPart.addStyledParagraphOfText("Heading2", 
            (table.isView() ? "View:" : "Table: ") + table.getName());

        final String comments = table.getComments();

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Description");
        mainDocumentPart.addParagraphOfText(comments == null ? "Comments" : comments);

        addTableColumns(table.getColumns());
    }

    private void addTableColumns(List<TableColumn> columns) {

        if (!CollectionUtils.isEmpty(columns)) {

            mainDocumentPart.addStyledParagraphOfText("Heading3", "Columns");

            Tbl tbl = createTable(TABLE_COLUMN_HEADERS, columns.size());

            // TblPr tblPr = tbl.getTblPr();         // create a brand new TblPr
            // TblStyle tblStyle = new TblStyle();          // create a brand new TblStyle
            // String styleID = 
            //         mainDocumentPart.getStyleDefinitionsPart().getIDForStyleName("Table Grid");      // find the style ID from the name
            
            // if (StringUtils.isNotEmpty(styleID)) {
            //     tblStyle.setVal(styleID);                    // just tell tblStyle what style it shall be
            //     tblPr.setTblStyle(tblStyle);                 // and affect each object its property...
            //     tbl.setTblPr(tblPr);
            // }

            //

            List<Object> rows = tbl.getContent();

            for (int row = 0; row < columns.size(); row++) {
                TableColumn tableColumn = columns.get(row);
                Tr tr = (Tr) rows.get(row + 1);

                List<Object> cells = tr.getContent();

                fillColumnNameCell(((Tc) cells.get(0)), tableColumn);

                setCellContent(cells, 1, tableColumn.getTypeName());
                setCellContent(cells, 2, tableColumn.getDetailedSize());
                setCellContent(cells, 3, Boolean.toString(tableColumn.isNullable()));

                if (tableColumn.isAutoUpdated()) {
                    setCellContent(cells, 4, Boolean.toString(tableColumn.isAutoUpdated()));
                }

                if (tableColumn.getDefaultValue() != null) {
                    setCellContent(cells, 5, "" + tableColumn.getDefaultValue());
                }
   
                setCellContent(cells, 6, tableColumn.getComments());
            }
        }
    }

    private void fillColumnNameCell(Tc cell, TableColumn tableColumn) {

        try {
            cell.getContent().clear();

            if (tableColumn.isPrimary()) {
                cell.getContent().add(createImageAndTextParagraph(getInlinePrimaryKeyImage(), tableColumn.getName()));
            }
            else if (tableColumn.isForeignKey()) {
                cell.getContent().add(createImageAndTextParagraph(getInlineForeignKeyImage(), tableColumn.getName()));
            }
            else {
                cell.getContent().add(
                    mainDocumentPart.createParagraphOfText(tableColumn.getName()));
            }
        }
        catch(Exception e) {
            LOGGER.error("fillColumnNameCell", e);
        }
    }


    private Inline getInlinePrimaryKeyImage() throws Exception {
        
        if (inlinePrimaryKeyImage == null) {
            InputStream is = getClass().getResourceAsStream(PK_IMAGE_PATH);
            byte[] bytes = IOUtils.toByteArray(is);
            
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, bytes);
            inlinePrimaryKeyImage = 
            imagePart.createImageInline("pkImage", "pkImage", 1, 2, false, 200);
        }

        return inlinePrimaryKeyImage;
    }

    private Inline getInlineForeignKeyImage() throws Exception {
        
        if (inlineForeignKeyImage == null) {
            InputStream is = getClass().getResourceAsStream(FK_IMAGE_PATH);
            byte[] bytes = IOUtils.toByteArray(is);
            
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordPackage, bytes);
            inlineForeignKeyImage = 
            imagePart.createImageInline("pkImage", "pkImage", 1, 2, false, 200);
        }

        return inlineForeignKeyImage;
    }

    private void addRelationships(Table table) {

        Collection<ForeignKeyConstraint> fks = table.getForeignKeys();

        if (CollectionUtils.isEmpty(fks)) {
            return;
        }

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Relationships");

        int size = 0;
        for (ForeignKeyConstraint foreignKeyConstraint : fks) {
            size += foreignKeyConstraint.getParentColumns().size();
        }

        Tbl tbl = createTable(FOREIGN_KEY_HEADERS, size);

        List<Object> rows = tbl.getContent();

        int rowIdx = 1;
        for (ForeignKeyConstraint foreignKeyConstraint : fks) {

            for (int i = 0; i < foreignKeyConstraint.getParentColumns().size(); i++) {
                Tr tr = (Tr) rows.get(rowIdx);
                List<Object> cells = tr.getContent();

                TableColumn parentColumn = foreignKeyConstraint.getParentColumns().get(i);
                TableColumn childColumn = foreignKeyConstraint.getChildColumns().get(i);

                setCellContent(cells, 0, foreignKeyConstraint.getName());
                setCellContent(cells, 1, childColumn.getName());
                setCellContent(cells, 2, parentColumn.getTable().getName() + "." + parentColumn.getName());
                setCellContent(cells, 3, foreignKeyConstraint.getDeleteRuleName());

                rowIdx++;
            }
        }
    }

    private void addChecks(Table table) {

        Map<String, String> checks = table.getCheckConstraints();

        if (CollectionUtils.isEmpty(checks)) {
            return;
        }

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Checks");

        Tbl tbl = createTable(CHECK_HEADERS, checks.size());

        List<Object> rows = tbl.getContent();

        int rowIdx = 0;

        for (Map.Entry<String, String> entry : checks.entrySet()) {
            Tr tr = (Tr) rows.get(++rowIdx);

            List<Object> cells = tr.getContent();
            
            setCellContent(cells, 0, entry.getKey());
            setCellContent(cells, 1, entry.getValue());
        }
    }

    private void addIndexes(Table table) {

        Collection<TableIndex> indexes = table.getIndexes();

        if (CollectionUtils.isEmpty(indexes)) {
            return;
        }

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Indexes");

        Tbl tbl = createTable(INDEXES_HEADERS, indexes.size());

        List<Object> rows = tbl.getContent();

        int rowIdx = 0;

        for (TableIndex index : indexes) {
            Tr tr = (Tr) rows.get(++rowIdx);

            List<Object> cells = tr.getContent();

            setCellContent(cells, 0, index.getName());
            setCellContent(cells, 1, index.getType());
            setCellContent(cells, 2, index.getColumnsAsString());
        }
    }

    /*
     * 
     * 
     */

    private void addView(View view) {

        mainDocumentPart.addStyledParagraphOfText("Heading2", "View:" + view.getName());

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Description");
        mainDocumentPart.addParagraphOfText(view.getComments());

        addViewColumns(view.getColumns());

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Source");

        createTable(new String [] { view.getViewDefinition() }, 0);
    }

    private void addViewColumns(List<TableColumn> columns) {

        if (!CollectionUtils.isEmpty(columns)) {

            mainDocumentPart.addStyledParagraphOfText("Heading3", "Columns");

            Tbl tbl = createTable(VIEW_COLUMN_HEADERS, columns.size());

            List<Object> rows = tbl.getContent();

            for (int row = 0; row < columns.size(); row++) {
                TableColumn tableColumn = columns.get(row);
                Tr tr = (Tr) rows.get(row + 1);

                List<Object> cells = tr.getContent();

                setCellContent(cells, 0, tableColumn.getName());
                setCellContent(cells, 1, tableColumn.getTypeName());
                setCellContent(cells, 2, tableColumn.getDetailedSize());
                setCellContent(cells, 3, Boolean.toString(tableColumn.isNullable()));
                setCellContent(cells, 4, tableColumn.getComments());
            }
        }
    }


    /*
     * 
     * 
     */
    private void addRoutinesList(Collection<Routine> routines) {

        if (CollectionUtils.isEmpty(routines)) {
            return;
        }

        mainDocumentPart.addStyledParagraphOfText("Heading2", "Routines");

        Tbl tbl = createTable(ROUTINES_HEADERS, routines.size());

        List<Object> rows = tbl.getContent();

        int rowIdx = 0;

        for (Routine routine : routines) {
            Tr tr = (Tr) rows.get(++rowIdx);

            List<Object> cells = tr.getContent();

            setCellContent(cells, 0, routine.getName());
            setCellContent(cells, 1, routine.getType());
            setCellContent(cells, 2, routine.getDefinitionLanguage());
            setCellContent(cells, 3, Boolean.toString(routine.isDeterministic()));
            setCellContent(cells, 4, routine.getReturnType());
            setCellContent(cells, 5, routine.getSecurityType());
            setCellContent(cells, 6, routine.getComment());
        }
    }


    private void addRoutine(Routine routine) {

        mainDocumentPart.addStyledParagraphOfText("Heading2", "Routine: " + routine.getName());

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Description");
        mainDocumentPart.addParagraphOfText(routine.getComment());

        addRoutineParameters(routine.getParameters());

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Source");

        createTable(new String [] {routine.getDefinition()}, 0);
    }

    private void addRoutineParameters(Collection<RoutineParameter> params) {

        if (CollectionUtils.isEmpty(params)) {
            return;
        }

        mainDocumentPart.addStyledParagraphOfText("Heading3", "Parameters");
        
        Tbl tbl = createTable(ROUTINE_PARAM_HEADERS, params.size());

        List<Object> rows = tbl.getContent();

        int rowIdx = 0;

        for (RoutineParameter param : params) {
            Tr tr = (Tr) rows.get(++rowIdx);

            List<Object> cells = tr.getContent();

            setCellContent(cells, 0, param.getName());
            setCellContent(cells, 1, param.getType());
            setCellContent(cells, 2, param.getMode());
        }
    }


    
    /*
     * Docx Utils 
     *  
     */
    private Tbl createTable(String [] columnTitles, int size) {

        int writableWidthTwips = 
        wordPackage.getDocumentModel().getSections().get(0)
                        .getPageDimensions().getWritableWidthTwips();

        Tbl tbl = 
            TblFactory.createTable(
                size + 1, 
                columnTitles.length, writableWidthTwips / columnTitles.length);

        List<Object> rows = tbl.getContent();

        List<Object> cells = ((Tr) rows.get(0)).getContent();

        for (int idx = 0; idx < columnTitles.length; idx++) {
            setCellContent(cells, idx, columnTitles[idx]);
        }

        mainDocumentPart.getContent().add(tbl);

        return tbl;
    }

    // private P createParagraphOfText(String str) {
    //     Text text = factory.createText();
    //     text.setValue(str); 

    //     R  run = factory.createR();
    //     run.getContent().add(text);

    //     P  p = factory.createP();
    //     p.getContent().add(run); 

    //     return p;
    // }

    public P createParagraphOfText(String simpleText) {
		return mainDocumentPart.createParagraphOfText(simpleText);
	}

    private void setCellContent(List<Object> cells, int cellIdx, String text) {
        Tc tc = (Tc) cells.get(cellIdx);

        tc.getContent().clear();

        tc.getContent().add(createParagraphOfText(text));
    }

    private P createImageAndTextParagraph(Inline inline, String text) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = factory.createP();
        R r = factory.createR();
        p.getContent().add(r);
        Drawing drawing = factory.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);
        
        Text t = factory.createText();
        t.setValue(text);
        r.getContent().add(t);
        return p;
    }

    private void addBr() {
        Br objBr = new Br();
            objBr.setType(STBrType.PAGE);

        mainDocumentPart.addObject(objBr);
    }


    // private WordprocessingMLPackage wordPackage;
    // private MainDocumentPart mainDocumentPart;

    // private void createAndRegisterTableStyle() {
    //     Style myNewStyle = Context.getWmlObjectFactory().createStyle();
    //     myNewStyle.setType("table");
    //     myNewStyle.setStyleId("myTableStyle");
        
    //     Style.Name n = Context.getWmlObjectFactory().createStyleName();
    //     n.setVal("myNewStyle");
    //     myNewStyle.setName(n);
    //     // Finally, add it to styles
    //     sdp.jaxbElement.getStyle().add(myNewStyle);         
        
    //     BasedOn based = Context.getWmlObjectFactory().createStyleBasedOn();
    //     based.setVal("TableGrid");      
    //     myNewStyle.setBasedOn(based);
    // }

    public void createPackage() throws InvalidFormatException {

        wordPackage = WordprocessingMLPackage.createPackage();
        mainDocumentPart = wordPackage.getMainDocumentPart();

        // Create a package
        // wordPackage = new WordprocessingMLPackage();
  
        // // Create main document part
        // mainDocumentPart = new MainDocumentPart();      
        
        // // Create main document part content
        // org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
        // org.docx4j.wml.Body  body = factory.createBody();      
        // org.docx4j.wml.Document wmlDocumentEl = factory.createDocument();
        
        // wmlDocumentEl.setBody(body);
        
        // // Create a basic sectPr using our Page model
        // PageDimensions page = new PageDimensions();
        // SectPr sectPr = factory.createSectPr();
        // body.setSectPr(sectPr);
        // sectPr.setPgSz(page.createPgSize() );
        // sectPr.setPgMar(page.createPgMar());
              
        // // Put the content in the part
        // mainDocumentPart.setJaxbElement(wmlDocumentEl);
                    
        // // Add the main document part to the package relationships
        // // (creating it if necessary)
        // wordPackage.addTargetPart(mainDocumentPart);
              
        // // Create a styles part
        // Part stylesPart = new org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart();
        // try {
        //    ((org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart) stylesPart)
        //          .unmarshalDefaultStyles();
           
        //    // Add the styles part to the main document part relationships
        //    // (creating it if necessary)
        //    mainDocumentPart.addTargetPart(stylesPart); // NB - add it to main doc part, not package!         
           
        // } catch (Exception e) {
        //    // TODO: handle exception
        //    e.printStackTrace();         
        // }
     }
}
