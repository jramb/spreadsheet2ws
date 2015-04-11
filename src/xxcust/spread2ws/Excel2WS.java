package xxcust.spread2ws;
/*
 * 2014 by J Ramb, Navigate Consulting
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.io.StringWriter;

import java.math.BigDecimal;

import java.util.Enumeration;
import java.util.Properties;
import java.util.UUID;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.soap.MessageFactory;
import javax.xml.soap.SOAPBody;
import javax.xml.soap.SOAPConnection;
import javax.xml.soap.SOAPConnectionFactory;
import javax.xml.soap.SOAPEnvelope;
import javax.xml.soap.SOAPHeader;
import javax.xml.soap.SOAPMessage;
import javax.xml.soap.SOAPPart;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMResult;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.transform.stream.StreamSource;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.Text;


public class Excel2WS {
    public Excel2WS() {
        super();
    }

    public static void myAssert(boolean cond, String message) {
        //assert(cond);
        if (!cond) {
            System.out.println(message);
            System.exit(1);
        }
    }

    public static void streamDOMSource(Source ds,
                                       StreamResult sr) throws TransformerException {
        //TransformerFactory tf = TransformerFactory.newInstance();
        Transformer transformer =
            TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
        transformer.setOutputProperty(OutputKeys.METHOD, "xml");
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount",
                                      "4");
        transformer.transform(ds, sr);
    }

    public static void printXMLDocument(Document doc,
                                        OutputStream out) throws IOException,
                                                                 TransformerException {
        streamDOMSource(new DOMSource(doc),
                        new StreamResult(new OutputStreamWriter(out,
                                                                "UTF-8")));
    }

    private static void printSOAPXML(SOAPMessage soapResponse,
                                     PrintStream out) throws Exception {
        streamDOMSource(soapResponse.getSOAPPart().getContent(), new StreamResult(out));
    }


    public static String xmlToString(Document doc) throws TransformerException {
        StringWriter sw = new StringWriter();
        streamDOMSource(new DOMSource(doc), new StreamResult(sw));
        return sw.toString(); //return sr.getWriter().toString();
    }


    // Oracles NVL
    public static String nvl(String val, String ifnull) {
        return (val == null || "".equals(val.trim())) ? ifnull : val;
    }

    public static XSSFCell setOrCreateCell(XSSFRow r, int col, String val) {
        XSSFCell c = r.getCell(col);
        if (c == null) {
            c = r.createCell(col);
        }
        c.setCellValue(val);
        if(col==0) {
            System.out.println(val);
        }
        return c;
    }

    public static String getCell(XSSFRow r, int col) {
        XSSFCell c = r.getCell(col);
        if (c == null)
            return null;
        switch (c.getCellType()) {
        case Cell.CELL_TYPE_NUMERIC:
            return (new BigDecimal(c.getNumericCellValue())).toString();
        case Cell.CELL_TYPE_BOOLEAN:
            return c.getBooleanCellValue()?"true":"false";
        case Cell.CELL_TYPE_STRING:
            return c.getStringCellValue();
            //break;
        default:
            return c.getStringCellValue();
            //break;
        }
    }

    static Document buildRowDoc(DocumentBuilder docBuilder, XSSFRow r, Properties prop) {
        Document doc = docBuilder.newDocument();
        Element rootElement = doc.createElement("root");
        doc.appendChild(rootElement);
        //rootElement.appendChild(doc.importNode(configXML,true)); // IMPORT the alien XML... DOES NOT WORK??
        // ok, doing it by hand {{{
        Element properties = doc.createElement("properties");
        rootElement.appendChild(properties);
        for (Enumeration e = prop.propertyNames(); e.hasMoreElements();
        ) {
            //for(Enumeration e: prop.propertyNames()) {
            String s = (String)e.nextElement();
            Element ent = doc.createElement("entry");
            ent.setAttribute("key", s);
            ent.setTextContent(prop.getProperty(s));
            properties.appendChild(ent);
        }
        // ok, doing it by hand }}}

        Element rowElement = doc.createElement("row");
        rootElement.appendChild(rowElement);

        for (int l = 0; l <= r.getLastCellNum(); l++) {
            String c = getCell(r, l); // nvl(getCell(r,l),"");
            if (c != null) {
                Element cell = doc.createElement("c");
                Attr attr = doc.createAttribute("col");
                attr.setValue(Integer.toString(l));
                cell.setAttributeNode(attr);
                Text txt = doc.createTextNode(c);
                cell.appendChild(txt);
                rowElement.appendChild(cell);
            }
        }

        return doc;    
    }
    


    public static void processWorksheet(XSSFSheet sheet, Properties prop) throws Exception {
        TransformerFactory transFact = TransformerFactory.newInstance();
        ClassLoader classLoader = Excel2WS.class.getClassLoader();
        // Transformer bodyTransform =   transFact.newTransformer(new StreamSource(prop.getProperty("body-transform")));
        Transformer infoTransform   = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("info-transform"))));
        Transformer bodyTransform   = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("body-transform"))));
        Transformer headerTransform = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("header-transform"))));
        Transformer resultTransform = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("result-transform"))));
        MessageFactory messageFactory =
            MessageFactory.newInstance(); // for SOAP calls

        boolean isDebug = "true".equals(prop.getProperty("debug"));
        String debugFileName = prop.getProperty("debug-file");
        PrintStream debugOut;
        if (debugFileName != null) {
            debugOut = new PrintStream(new File(debugFileName));
        } else {
            debugOut = System.out;
        }

        String ep;
        String env = prop.getProperty("environment");
        debugOut.println("env="+env);
        ep = nvl(prop.getProperty("endpoint-"+env), prop.getProperty("endpoint"));

        myAssert(ep != null, "Config: environment and/or endpoint must be defined empty");
        debugOut.println("Endpoint: " + ep);

        DocumentBuilderFactory docFactory =
            DocumentBuilderFactory.newInstance();
        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();

        // Create SOAP Connection
        SOAPConnectionFactory soapConnectionFactory =
            SOAPConnectionFactory.newInstance();
        SOAPConnection soapConnection =
            soapConnectionFactory.createConnection();


        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow r = sheet.getRow(i);
            //System.out.println("Processing: " + r.getCell(2));
            
            prop.setProperty("rownum", Integer.toString(i));
            Document doc = buildRowDoc(docBuilder, r, prop);
            StringWriter inf = new StringWriter();
            //printXMLDocument(doc, System.out);
            infoTransform.transform(new DOMSource(doc), new StreamResult(inf));
            String infoStr = inf.toString();

            if (!"".equals(infoStr)) {
                System.out.print("Row "+(i+1)+": "+infoStr+": ");
                //System.out.print("Processing: " + r.getCell(2) + " = ");
                debugOut.println((i+1)+": "+infoStr);
                


                if (isDebug) {
                    //System.out.println(xmlToString(doc));
                    printXMLDocument(doc, debugOut);
                    DOMResult dr = new DOMResult();
                    // StreamResult sr =  new StreamResult(new OutputStreamWriter(System.out, "UTF-8"));
                    bodyTransform.transform(new DOMSource(doc), dr);
                    printXMLDocument((Document)dr.getNode(), debugOut);
                }

                SOAPMessage soapMessage = messageFactory.createMessage();
                SOAPPart soapPart = soapMessage.getSOAPPart();
                SOAPEnvelope envelope = soapPart.getEnvelope();

                // SOAP Envelope
                SOAPHeader soapHdr = envelope.getHeader();
                SOAPBody soapBody = envelope.getBody();

                // SOAP Body
                bodyTransform.transform(new DOMSource(doc),
                                        new DOMResult(soapBody));
                headerTransform.transform(new DOMSource(doc),
                                          new DOMResult(soapHdr));
                if (isDebug) {
                    printSOAPXML(soapMessage, debugOut);
                }

                // Done
                soapMessage.saveChanges();
                SOAPMessage soapResponse = null;
                try {
                    soapResponse = soapConnection.call(soapMessage, ep);
                    if (isDebug) {
                        printSOAPXML(soapResponse, debugOut);
                        //printXMLDocument(soapResponse.getSOAPPart().getContent(),System.out);
                    }
                    DOMResult resDom = new DOMResult();
                    resultTransform.transform(soapResponse.getSOAPPart().getContent(),
                                              resDom);
                    Document res = (Document)resDom.getNode();
                    if (isDebug) {
                        printXMLDocument(res, debugOut);
                    }
                    //setOrCreateCell(r, 0, "ok");
                    //printXMLDocument(res, System.out);

                    if (res != null && res.getDocumentElement() != null) {
                        NodeList cols =
                            res.getDocumentElement().getChildNodes();
                        for (int k = 0; k < cols.getLength(); k++) {
                            Node cx = cols.item(k);
                            int colnr =
                                Integer.parseInt(cx.getAttributes().getNamedItem("col").getTextContent());
                            setOrCreateCell(r, colnr, cx.getTextContent());
                        }
                    } else {
                        setOrCreateCell(r, 0, "ERROR");
                        setOrCreateCell(r, 1, "Result after transformation is empty, check logs");
                    }
                } catch (Exception e) {
                    e.printStackTrace(System.out);
                    System.out.println(e.getLocalizedMessage());
                    setOrCreateCell(r, 0, "ERROR");
                    setOrCreateCell(r, 1, e.getLocalizedMessage());
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        //ExcelImport excelImport = new ExcelImport();
        System.out.println("excelwsload by jramb");
        String filename;
        Properties prop = new Properties();
        ClassLoader classLoader = Excel2WS.class.getClassLoader();
        InputStream config =
            classLoader.getResourceAsStream("config.xml"); // config.properties?

        myAssert(config != null,
                 "Could not find config.xml in path, I need that!");
        //prop.load(config);  // if config.properties
        prop.loadFromXML(config);
        //properties.load(Class.getClassLoader().getResourceAsStream("config.properties"));

        //config = classLoader.getResourceAsStream("config.xml"); // reopen necessary
        //Document confXML =  DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(config);
        //printXMLDocument(confXML, System.out);
        prop.setProperty("now",
                         new java.text.SimpleDateFormat("yyyy-MM-dd").format(new java.util.Date()));
        prop.setProperty("uuid", UUID.randomUUID().toString());

        int i=0;
        if(args.length>0) {
            File f = new File(args[i]);
            if (f.exists()) {
                prop.setProperty("excelfile", f.getName());
                i++;
            }
        }
        for (/*int i = 0*/; i < args.length; i++) {
            String[] v = args[i].split("=");
            if (v.length == 2) {
                prop.setProperty(v[0], v[1]);
            }
        }
        filename = prop.getProperty("excelfile");
        System.out.println("Starting spreadsheet processing: " + filename);
        File xlsx = new File(filename);
        myAssert(xlsx.canRead(), "File not found: " + xlsx.getName());
        myAssert(xlsx.canWrite(), "File not writeable: " + xlsx.getName());

        XSSFWorkbook x = new XSSFWorkbook(new FileInputStream(xlsx));

        loadOverrideProperties(x.getSheetAt(1), prop);

        // make sure we can write the file by... writing it
        boolean updateFile = !"false".equals(prop.getProperty("update-file"));
    
        try {
            if (updateFile) {
                x.write(new FileOutputStream(xlsx));
            }
        } catch (IOException e) {
            System.out.println("Can not write to " + filename);
            System.out.println("Maybe you need to close the file in Excel?");
            System.exit(1);
        }
        // now process the file
        try {
            processWorksheet(x.getSheetAt(0), prop);
        } finally {
            if (updateFile) {
                System.out.println("Updating " + xlsx.getName());
                x.write(new FileOutputStream(xlsx));
            }
        }
        System.out.println("Done.");
    }

    private static void loadOverrideProperties(XSSFSheet sheet,
                                               Properties prop) {
        if(sheet == null) {
            return;
        }
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow r = sheet.getRow(i);
            if (r != null) {
                String k = getCell(r, 0);
                String v = getCell(r, 1);
                if(!"".equals(k)) {
                    prop.setProperty(k, v);
                    //System.out.println(k+" = "+v);
                }
            }
        }
    }
}
