package com.randominteger.examples.aspose;

import com.aspose.words.*;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.Charset;

/**
 * @author Olavi Tiimus
 */
public class Main {

  private static final Charset CHARSET = Charset.forName("windows-1252");

  private void generate() throws Exception {
    Document appointmentsDocument = new Document(new ByteArrayInputStream(getRtfMarkup().getBytes(CHARSET)));
    Document destinationDocument = new Document(getTemplateFile());

    Node insertAfterNode = destinationDocument.getFirstSection().getBody().getFirstParagraph();
    NodeImporter importer = new NodeImporter(appointmentsDocument, insertAfterNode.getDocument(), ImportFormatMode.USE_DESTINATION_STYLES);
    CompositeNode dstStory = insertAfterNode.getParentNode();

    if (appointmentsDocument.hasChildNodes()) {
      for (Node node : (NodeCollection<Node>) appointmentsDocument.getFirstSection().getBody().getChildNodes()) {
        if (node.getNodeType() == NodeType.PARAGRAPH) {
          Node importedNode = importer.importNode(node, true);

          dstStory.insertAfter(importedNode, insertAfterNode);
          insertAfterNode = importedNode;
        }
      }
    }

    destinationDocument.save("result.rtf", SaveFormat.RTF);
  }

  private String getRtfMarkup() {
    StringBuilder rtfMarkup = new StringBuilder();

    // header
    rtfMarkup.append("{\\rtf1\\ansi\\deff0");
    rtfMarkup.append("{\\stylesheet ");
    rtfMarkup.append("{\\s2 Kirche;}");
    rtfMarkup.append("{\\s3 Tag;}");
    rtfMarkup.append("{\\s4 Termin;}");
    rtfMarkup.append("}");

    // 1st day
    rtfMarkup.append("\\pard\\plain\\s3\\fs22\\i\\b Montag\\tab 01.06.\\tab Hl. Justin \\i0\\b0\\par");
    rtfMarkup.append("\\pard\\plain\\s4\\tab \\tab 20.00\\tab Taizè-Gebet im Pfarrheim\\i \\i0\\par");

    // 2nd day
    rtfMarkup.append("\\pard\\plain\\s3\\fs22\\i\\b Dienstag\\tab 02.06.\\tab Hl. Marcellinus und hl. Petrus \\i0\\b0\\par");
    rtfMarkup.append("\\pard\\plain\\s4\\tab\\tab 19.00\\tab Messfeier\\i (Pfarrvikar Norbert Becker)\\i0\\par");
    rtfMarkup.append("\\pard\\plain\\s4\\tab\\tab 19.00\\tab Messfeier  u. Gebet um geistliche Berufe mit Kollekte für Achacachi mit kurzer Anbetung u. Aussetzung Elfriede und Maria Stemmer und Angehörige\\i (Pfarrvikar Norbert Becker)\\i0\\par");
    rtfMarkup.append("\\pard\\plain\\s4\\tab\\tab 20.00\\tab Gemeinsames Gebet im Pfarrheim\\i \\i0\\par");

    // footer
    rtfMarkup.append("}");

    return rtfMarkup.toString();
  }

  private InputStream getTemplateFile() {
    return getClass().getClassLoader().getResourceAsStream("template.rtf");
  }

  public static void main(String[] args) {
    try {
      new Main().generate();
    }
    catch (Exception e) {
      e.printStackTrace();
    }
  }
}
