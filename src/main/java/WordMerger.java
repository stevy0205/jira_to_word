import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;

import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.math.BigInteger;

public class WordMerger {

    private Map<BigInteger, BigInteger> numIDs = null; // to handle merging numID

    public WordMerger() {
        this.numIDs= new HashMap<BigInteger, BigInteger>();
    }

    private void traverseBodyElements(List<IBodyElement> bodyElements, IBody resultBody) throws Exception {
        for (IBodyElement bodyElement : bodyElements) {
            if (bodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = (XWPFParagraph)bodyElement;
                XWPFParagraph resultParagraph = createParagraphWithPPr(paragraph, resultBody);
                traverseRunElements(paragraph.getIRuns(), resultParagraph);
            } else if (bodyElement instanceof XWPFSDT) {
                XWPFSDT sDT = (XWPFSDT)bodyElement;
                XWPFSDT resultSDT = createSDT(sDT, resultBody);
                //ToDo: handle further issues ...
            } else if (bodyElement instanceof XWPFTable) {
                XWPFTable table = (XWPFTable)bodyElement;
                XWPFTable resultTable = createTableWithTblPrAndTblGrid(table, resultBody);
                traverseTableRows(table.getRows(), resultTable);
            }
        }
    }

    private XWPFSDT createSDT(XWPFSDT sDT, IBody resultBody) {
        //not ready yet
        //we simply add paragraphs to avoid corruped documents
        if (resultBody instanceof XWPFDocument) {
            XWPFDocument resultDocument = (XWPFDocument)resultBody;
            XWPFParagraph resultParagraph = resultDocument.createParagraph();
            //ToDo: handle further issues ...
        } else if (resultBody instanceof XWPFTableCell) {
            XWPFTableCell resultTableCell = (XWPFTableCell)resultBody;
            XWPFParagraph resultParagraph = resultTableCell.addParagraph();
            //ToDo: handle further issues ...
        } //ToDo: else others ...
        //ToDo: handle SDT properly
        return null;
    }

    private XWPFParagraph createParagraphWithPPr(XWPFParagraph paragraph, IBody resultBody) {
        if (resultBody instanceof XWPFDocument) {
            XWPFDocument resultDocument = (XWPFDocument)resultBody;
            XWPFParagraph resultParagraph = resultDocument.createParagraph();
            resultParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());//simply copy the underlying XML bean to avoid more code
            handleStyles(resultDocument, paragraph);
            handleNumberings(paragraph, resultParagraph);
            //ToDo: handle further issues ...
            return resultParagraph;
        } else if (resultBody instanceof XWPFTableCell) {
            XWPFTableCell resultTableCell = (XWPFTableCell)resultBody;
            XWPFParagraph resultParagraph = resultTableCell.addParagraph();
            resultParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());//simply copy the underlying XML bean to avoid more code
            handleStyles(resultTableCell, paragraph);
            //ToDo: handle further issues ...
            return resultParagraph;
        } //ToDo: else others ...
        return null;
    }

    private void handleNumberings(XWPFParagraph paragraph, XWPFParagraph resultParagraph) {
        //if we have numberings, we need merging the numIDs and abstract numberings of the two different documents
        BigInteger numID = paragraph.getNumID();
        if (numID == null) return;
        BigInteger resultNumID = this.numIDs.get(numID);
        if (resultNumID == null) {
            XWPFDocument document = paragraph.getDocument();
            XWPFNumbering numbering = document.createNumbering();
            XWPFNum num = numbering.getNum(numID);
            BigInteger abstractNumID = numbering.getAbstractNumID(numID);
            XWPFAbstractNum abstractNum = numbering.getAbstractNum(abstractNumID);
            XWPFAbstractNum resultAbstractNum = new XWPFAbstractNum((org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum)abstractNum.getCTAbstractNum().copy());
            XWPFDocument resultDocument = resultParagraph.getDocument();
            XWPFNumbering resultNumbering = resultDocument.createNumbering();
            int pos = resultNumbering.getAbstractNums().size();
            resultAbstractNum.getCTAbstractNum().setAbstractNumId(BigInteger.valueOf(pos));
            BigInteger resultAbstractNumID = resultNumbering.addAbstractNum(resultAbstractNum);
            resultNumID = resultNumbering.addNum(resultAbstractNumID);
            XWPFNum resultNum = resultNumbering.getNum(resultNumID);
            resultNum.getCTNum().setLvlOverrideArray(num.getCTNum().getLvlOverrideArray());
            this.numIDs.put(numID, resultNumID);
        }
        resultParagraph.setNumID(resultNumID);
    }

    private void handleStyles(IBody resultBody, IBodyElement bodyElement) {
        //if we have bodyElement styles we need merging those styles of the two different documents
        XWPFDocument document = null;
        String styleID = null;
        if (bodyElement instanceof XWPFParagraph) {
            XWPFParagraph paragraph = (XWPFParagraph)bodyElement;
            document = paragraph.getDocument();
            styleID = paragraph.getStyleID();
        } else if (bodyElement instanceof XWPFTable) {
            XWPFTable table = (XWPFTable)bodyElement;
            if (table.getPart() instanceof XWPFDocument) {
                document = (XWPFDocument)table.getPart();
                styleID = table.getStyleID();
            }
        } //ToDo: else others ...
        if (document == null || styleID == null || "".equals(styleID)) return;
        XWPFDocument resultDocument = null;
        if (resultBody instanceof XWPFDocument) {
            resultDocument = (XWPFDocument)resultBody;
        } else if (resultBody instanceof XWPFTableCell) {
            XWPFTableCell resultTableCell = (XWPFTableCell)resultBody;
            resultDocument = resultTableCell.getXWPFDocument();
        } //ToDo: else others ...
        if (resultDocument != null) {
            XWPFStyles styles = document.getStyles();
            XWPFStyles resultStyles = resultDocument.getStyles();
            XWPFStyle style = styles.getStyle(styleID);
            //merge each used styles, also the related ones
            for (XWPFStyle relatedStyle : styles.getUsedStyleList(style)) {
                if (resultStyles.getStyle(relatedStyle.getStyleId()) == null) {
                    resultStyles.addStyle(relatedStyle);
                }
            }
        }
    }

    private XWPFTable createTableWithTblPrAndTblGrid(XWPFTable table, IBody resultBody) {
        if (resultBody instanceof XWPFDocument) {
            XWPFDocument resultDocument = (XWPFDocument)resultBody;
            XWPFTable resultTable = resultDocument.createTable();
            resultTable.removeRow(0);
            resultTable.getCTTbl().setTblPr(table.getCTTbl().getTblPr());//simply copy the underlying XML bean to avoid more code
            resultTable.getCTTbl().setTblGrid(table.getCTTbl().getTblGrid());//simply copy the underlying XML bean to avoid more code
            handleStyles(resultDocument, table);
            //ToDo: handle further issues ...
            return resultTable;
        } else if (resultBody instanceof XWPFTableCell) {
            //ToDo: handle stacked tables
        } //ToDo: else others ...
        return null;
    }

    private void traverseRunElements(List<IRunElement> runElements, IRunBody resultRunBody) throws Exception {
        for (IRunElement runElement : runElements) {
            if (runElement instanceof XWPFFieldRun) {
                XWPFFieldRun fieldRun = (XWPFFieldRun)runElement;
                XWPFFieldRun resultFieldRun = createFieldRunWithRPr(fieldRun, resultRunBody);
                traversePictures(fieldRun, resultFieldRun);
            } else if (runElement instanceof XWPFHyperlinkRun) {
                XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun)runElement;
                XWPFHyperlinkRun resultHyperlinkRun = createHyperlinkRunWithRPr(hyperlinkRun, resultRunBody);
                traversePictures(hyperlinkRun, resultHyperlinkRun);
            } else if (runElement instanceof XWPFRun) {
                XWPFRun run = (XWPFRun)runElement;
                XWPFRun resultRun = createRunWithRPr(run, resultRunBody);
                traversePictures(run, resultRun);
            } else if (runElement instanceof XWPFSDT) {
                XWPFSDT sDT = (XWPFSDT)runElement;
                //ToDo: handle SDT
            }
        }
    }

    private void copyTextOfRuns(XWPFRun run, XWPFRun resultRun) {
        //copy all of the possible T contents of the runs
        for (int i = 0; i < run.getCTR().sizeOfTArray(); i++) {
            resultRun.setText(run.getText(i), i);
        }
    }

    private XWPFFieldRun createFieldRunWithRPr(XWPFFieldRun fieldRun, IRunBody resultRunBody) {
        if (resultRunBody instanceof XWPFParagraph) {
            XWPFParagraph resultParagraph = (XWPFParagraph)resultRunBody;
            XWPFFieldRun resultFieldRun = (XWPFFieldRun)resultParagraph.createRun();
            resultFieldRun.getCTR().setRPr(fieldRun.getCTR().getRPr());//simply copy the underlying XML bean to avoid more code
            //ToDo: handle field runs properly ...
            handleRunStyles(resultParagraph.getDocument(), fieldRun);
            //ToDo: handle further issues ...
            return resultFieldRun;
        } else if (resultRunBody instanceof XWPFSDT) {
            //ToDo: handle SDT
        }
        return null;
    }

    private XWPFHyperlinkRun createHyperlinkRunWithRPr(XWPFHyperlinkRun hyperlinkRun, IRunBody resultRunBody) {
        if (resultRunBody instanceof XWPFParagraph) {
            XWPFParagraph resultParagraph = (XWPFParagraph)resultRunBody;
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink resultCTHyperLink = resultParagraph.getCTP().addNewHyperlink();
            resultCTHyperLink.addNewR();
            XWPFHyperlinkRun resultHyperlinkRun =  new XWPFHyperlinkRun(resultCTHyperLink, resultCTHyperLink.getRArray(0), resultParagraph);
            if (hyperlinkRun.getAnchor() != null) {
                resultHyperlinkRun = resultParagraph.createHyperlinkRun(hyperlinkRun.getAnchor());
            }
            resultHyperlinkRun.getCTR().setRPr(hyperlinkRun.getCTR().getRPr());//simply copy the underlying XML bean to avoid more code
            copyTextOfRuns(hyperlinkRun, resultHyperlinkRun);
            //ToDo: handle external hyperlink runs properly ...
            handleRunStyles(resultParagraph.getDocument(), hyperlinkRun);
            //ToDo: handle further issues ...
            return resultHyperlinkRun;
        } else if (resultRunBody instanceof XWPFSDT) {
            //ToDo: handle SDT
        }
        return null;
    }

    private XWPFRun createRunWithRPr(XWPFRun run, IRunBody resultRunBody) {
        if (resultRunBody instanceof XWPFParagraph) {
            XWPFParagraph resultParagraph = (XWPFParagraph)resultRunBody;
            XWPFRun resultRun = resultParagraph.createRun();
            resultRun.getCTR().setRPr(run.getCTR().getRPr());//simply copy the underlying XML bean to avoid more code
            copyTextOfRuns(run, resultRun);
            handleRunStyles(resultParagraph.getDocument(), run);
            //ToDo: handle further issues ...
            return resultRun;
        } else if (resultRunBody instanceof XWPFSDT) {
            //ToDo: handle SDT
        }
        return null;
    }

    private void handleRunStyles(IBody resultBody, IRunElement runElement) {
        //if we have runElement styles we need merging those styles of the two different documents
        XWPFDocument document = null;
        String styleID = null;
        if (runElement instanceof XWPFRun) {
            XWPFRun run = (XWPFRun)runElement;
            document = run.getDocument();
            styleID = run.getStyle();
        } else if (runElement instanceof XWPFHyperlinkRun) {
            XWPFHyperlinkRun run = (XWPFHyperlinkRun)runElement;
            document = run.getDocument();
            styleID = run.getStyle();
        } else if (runElement instanceof XWPFFieldRun) {
            XWPFFieldRun run = (XWPFFieldRun)runElement;
            document = run.getDocument();
            styleID = run.getStyle();
        } //ToDo: else others ...
        if (document == null || styleID == null || "".equals(styleID)) return;
        XWPFDocument resultDocument = null;
        if (resultBody instanceof XWPFDocument) {
            resultDocument = (XWPFDocument)resultBody;
        } else if (resultBody instanceof XWPFTableCell) {
            XWPFTableCell resultTableCell = (XWPFTableCell)resultBody;
            resultDocument = resultTableCell.getXWPFDocument();
        } //ToDo: else others ...
        if (resultDocument != null) {
            XWPFStyles styles = document.getStyles();
            XWPFStyles resultStyles = resultDocument.getStyles();
            XWPFStyle style = styles.getStyle(styleID);
            //merge each used styles, also the related ones
            for (XWPFStyle relatedStyle : styles.getUsedStyleList(style)) {
                if (resultStyles.getStyle(relatedStyle.getStyleId()) == null) {
                    resultStyles.addStyle(relatedStyle);
                }
            }
        }
    }

    private void traverseTableRows(List<XWPFTableRow> tableRows, XWPFTable resultTable) throws Exception {
        for (XWPFTableRow tableRow : tableRows) {
            XWPFTableRow resultTableRow = createTableRowWithTrPr(tableRow, resultTable);
            traverseTableCells(tableRow.getTableICells(), resultTableRow);
        }
    }

    private XWPFTableRow createTableRowWithTrPr(XWPFTableRow tableRow, XWPFTable resultTable) {
        XWPFTableRow resultTableRow = resultTable.createRow();
        for (int i = resultTableRow.getTableCells().size(); i > 0; i--) { //table row should be empty at first
            resultTableRow.removeCell(i-1);
        }
        resultTableRow.getCtRow().setTrPr(tableRow.getCtRow().getTrPr());//simply copy the underlying XML bean to avoid more code
        //ToDo: handle further issues ...
        return resultTableRow;
    }

    private void traverseTableCells(List<ICell> tableICells, XWPFTableRow resultTableRow) throws Exception {
        for (ICell tableICell : tableICells) {
            if (tableICell instanceof XWPFSDTCell) {
                XWPFSDTCell sDTCell = (XWPFSDTCell)tableICell;
                XWPFSDTCell resultSdtTableCell = createSdtTableCell(sDTCell, resultTableRow);
                //ToDo: handle further issues ...
            } else if (tableICell instanceof XWPFTableCell) {
                XWPFTableCell tableCell = (XWPFTableCell)tableICell;
                XWPFTableCell resultTableCell = createTableCellWithTcPr(tableCell, resultTableRow);
                traverseBodyElements(tableCell.getBodyElements(), resultTableCell);
            }
        }
    }

    private XWPFSDTCell createSdtTableCell(XWPFSDTCell sDTCell, XWPFTableRow resultTableRow) {
        //create at least a cell to avoid corrupted document
        XWPFTableCell resultTableCell = resultTableRow.createCell();
        //ToDo: handle SDTCell properly
        //ToDo: handle further issues ...
        return null;
    }

    private XWPFTableCell createTableCellWithTcPr(XWPFTableCell tableCell, XWPFTableRow resultTableRow) {
        XWPFTableCell resultTableCell = resultTableRow.createCell();
        resultTableCell.removeParagraph(0);
        resultTableCell.getCTTc().setTcPr(tableCell.getCTTc().getTcPr());//simply copy the underlying XML bean to avoid more code
        //ToDo: handle further issues ...
        return resultTableCell;
    }

    private void traversePictures(IRunElement runElement, IRunElement resultRunElement) throws Exception {
        List<XWPFPicture> pictures = null;
        if (runElement instanceof XWPFFieldRun) {
            XWPFFieldRun fieldRun = (XWPFFieldRun)runElement;
            pictures = fieldRun.getEmbeddedPictures();
        } else if (runElement instanceof XWPFHyperlinkRun) {
            XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun)resultRunElement;
            pictures = hyperlinkRun.getEmbeddedPictures();
        } else if (runElement instanceof XWPFRun) {
            XWPFRun run = (XWPFRun)runElement;
            pictures = run.getEmbeddedPictures();
        } else if (runElement instanceof XWPFSDT) {
            XWPFSDT sDT = (XWPFSDT)runElement;
            //ToDo: handle SDT
        }
        if (pictures != null) {
            for (XWPFPicture picture : pictures) {
                XWPFPictureData pictureData = picture.getPictureData();
                XWPFPicture resultPicture = createPictureWithDrawing(runElement, picture, pictureData, resultRunElement);
            }
        }
    }

    private XWPFPicture createPictureWithDrawing(IRunElement runElement, XWPFPicture picture, XWPFPictureData pictureData, IRunElement resultRunElement) {
        if (resultRunElement instanceof XWPFFieldRun) {
            XWPFFieldRun fieldRun = (XWPFFieldRun)runElement;
            XWPFFieldRun resultFieldRun = (XWPFFieldRun)resultRunElement;
            XWPFPicture resultPicture = createPictureWithDrawing(fieldRun, resultFieldRun, picture, pictureData);
            return resultPicture;
        } else if (resultRunElement instanceof XWPFHyperlinkRun) {
            XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun)runElement;
            XWPFHyperlinkRun resultHyperlinkRun = (XWPFHyperlinkRun)resultRunElement;
            XWPFPicture resultPicture = createPictureWithDrawing(hyperlinkRun, resultHyperlinkRun, picture, pictureData);
            return resultPicture;
        } else if (resultRunElement instanceof XWPFRun) {
            XWPFRun run = (XWPFRun)runElement;
            XWPFRun resultRun = (XWPFRun)resultRunElement;
            XWPFPicture resultPicture = createPictureWithDrawing(run, resultRun, picture, pictureData);
            return resultPicture;
        } else if (resultRunElement instanceof XWPFSDT) {
            XWPFSDT sDT = (XWPFSDT)resultRunElement;
            //ToDo: handle SDT
        }
        return null;
    }

    private XWPFPicture createPictureWithDrawing(XWPFRun run, XWPFRun resultRun, XWPFPicture picture, XWPFPictureData pictureData) {
        try {
            XWPFPicture resultPicture = resultRun.addPicture(
                    pictureData.getPackagePart().getInputStream(),
                    pictureData.getPictureType(),
                    pictureData.getFileName(),
                    Units.pixelToEMU((int)picture.getWidth()),
                    Units.pixelToEMU((int)picture.getDepth()));
            String rId = resultPicture.getCTPicture().getBlipFill().getBlip().getEmbed();
            resultRun.getCTR().setDrawingArray(0, run.getCTR().getDrawingArray(0));//simply copy the underlying XML bean to avoid more code
            //but then correct the rID
            String declareNameSpaces = "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main'; ";
            org.apache.xmlbeans.XmlObject[] selectedObjects = resultRun.getCTR().getDrawingArray(0).selectPath(
                    declareNameSpaces
                            + "$this//a:blip");
            for (org.apache.xmlbeans.XmlObject blipObject : selectedObjects) {
                if (blipObject instanceof org.openxmlformats.schemas.drawingml.x2006.main.CTBlip) {
                    org.openxmlformats.schemas.drawingml.x2006.main.CTBlip blip = (org.openxmlformats.schemas.drawingml.x2006.main.CTBlip)blipObject;
                    if (blip.isSetEmbed()) blip.setEmbed(rId);
                }
            }
            //remove rIDs to external hyperlinks to avoid corruot document
            selectedObjects = resultRun.getCTR().getDrawingArray(0).selectPath(
                    declareNameSpaces
                            + "$this//a:hlinkClick");
            for (org.apache.xmlbeans.XmlObject hlinkClickObject : selectedObjects) {
                if (hlinkClickObject instanceof org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink) {
                    org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink hlinkClick = (org.openxmlformats.schemas.drawingml.x2006.main.CTHyperlink)hlinkClickObject;
                    if (hlinkClick.isSetId()) hlinkClick.setId("");
                    //ToDo: handle pictures having hyperlinks properly
                }
            }
            //ToDo: handle further issues ...
            return resultPicture;
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }

    public void merge(String firstFilePath, String secondFilePath, String resultFilePath) throws Exception {
        XWPFDocument resultDocument = new XWPFDocument(new FileInputStream(firstFilePath));
        XWPFDocument documentToAppend = new XWPFDocument(new FileInputStream(secondFilePath));
        traverseBodyElements(documentToAppend.getBodyElements(), resultDocument);
        documentToAppend.close();
        FileOutputStream out = new FileOutputStream(resultFilePath);
        resultDocument.write(out);
        out.close();
        resultDocument.close();
    }
}