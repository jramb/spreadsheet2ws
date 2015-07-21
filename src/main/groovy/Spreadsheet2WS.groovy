/*
 * 2014-2015 by J Ramb, Navigate Consulting
 */
import groovy.transform.CompileStatic
//import groovy.transform.TypeChecked

import javax.xml.parsers.DocumentBuilder
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.soap.MessageFactory
import javax.xml.soap.SOAPBody
import javax.xml.soap.SOAPConnection
import javax.xml.soap.SOAPConnectionFactory
import javax.xml.soap.SOAPEnvelope
import javax.xml.soap.SOAPHeader
import javax.xml.soap.SOAPMessage
import javax.xml.soap.SOAPPart
import javax.xml.transform.OutputKeys
import javax.xml.transform.Source
import javax.xml.transform.Transformer
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMResult
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult
import javax.xml.transform.stream.StreamSource

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import org.w3c.dom.Attr
import org.w3c.dom.Document
import org.w3c.dom.Element
import org.w3c.dom.Node
import org.w3c.dom.NodeList
import org.w3c.dom.Text


static void myAssert(boolean cond, String message) {
    if (!cond) {
        println message
        System.exit(1)
    }
}


static void streamDOMSource(Source ds, StreamResult sr) {
    Transformer transformer =
            TransformerFactory.newInstance().newTransformer()
    transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes")
    transformer.setOutputProperty(OutputKeys.METHOD, "xml")
    transformer.setOutputProperty(OutputKeys.INDENT, "yes")
    transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8")
    transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4")

    transformer.transform(ds, sr)
}

static void printXMLDocument(Document doc, OutputStream out) {
    streamDOMSource(new DOMSource(doc),
            new StreamResult(new OutputStreamWriter(out, "UTF-8")))
}

static void printSOAPXML(SOAPMessage soapResponse, PrintStream out) {
    streamDOMSource(soapResponse.getSOAPPart().getContent(), new StreamResult(out))
}


static String xmlToString(Document doc) {
    StringWriter sw = new StringWriter()
    streamDOMSource(new DOMSource(doc), new StreamResult(sw))
    return sw.toString()
}


// this gets problems unless compiled static (createCell)
@CompileStatic
static XSSFCell setOrCreateCell(XSSFRow r, int col, String val) {
    XSSFCell c = r.getCell(col)
    if (c == null) {
        c = r.createCell(col)
    }
    c.setCellValue(val)
    if (col == 0) {
        System.out.println(val)
    }
    return c
}


// this gets problems unless compiled static (getCellType)
@CompileStatic
static String getCell(XSSFRow r, int col) {
    XSSFCell c = r.getCell(col)
    if (c == null) {
        return null
    }
    switch (c.getCellType()) {
        case Cell.CELL_TYPE_NUMERIC:
            return new BigDecimal(c.getNumericCellValue()).toString()
        case Cell.CELL_TYPE_BOOLEAN:
            return c.getBooleanCellValue() ? "true" : "false"
        case Cell.CELL_TYPE_STRING:
            return c.getStringCellValue()
        case Cell.CELL_TYPE_FORMULA:
            return c.getCellFormula()
        default:
            return c.getStringCellValue()
    }
}


static Document buildRowDoc(DocumentBuilder docBuilder, XSSFRow r, Properties prop) {
    Document doc = docBuilder.newDocument()
    Element rootElement = doc.createElement("root")
    doc.appendChild(rootElement)
    //rootElement.appendChild(doc.importNode(configXML,true)); // IMPORT the alien XML... DOES NOT WORK??
    // ok, doing it by hand {{{
    Element properties = doc.createElement("properties")
    rootElement.appendChild(properties)
    for (Enumeration e = prop.propertyNames(); e.hasMoreElements();) {
    //(prop.propertyNames() as Enumeration).each { e ->
        String s = (String) e.nextElement()
        Element ent = doc.createElement("entry")
        ent.setAttribute("key", s)
        ent.setTextContent(prop."$s")
        properties.appendChild(ent)
    }
    // ok, doing it by hand }}}

    Element rowElement = doc.createElement("row")
    rootElement.appendChild(rowElement)

    for (int l = 0; l <= r.getLastCellNum(); l++) {
        String c = getCell(r, l)
        if (c != null) {
            Element cell = doc.createElement("c")
            Attr attr = doc.createAttribute("col")
            attr.setValue(Integer.toString(l))
            cell.setAttributeNode(attr)
            Text txt = doc.createTextNode(c)
            cell.appendChild(txt)
            rowElement.appendChild(cell)
        }
    }

    return doc
}

static Transformer loadTransformer(TransformerFactory transFact, Properties prop, String name) {
    def styleSheet = prop[prop[name]]
    //println "StyleSheet=$styleSheet"
    Transformer transform
    if (styleSheet == null) {
        //InputStream istr = this.getClass().classLoader.getResourceAsStream(prop[name]);
        InputStream istr = Spreadsheet2WS.classLoader.getResourceAsStream(prop[name])
        assert istr != null, "Could not find resource: ${prop[name]}"
        transform = transFact.newTransformer(new StreamSource(istr))
    } else {
        transform = transFact.newTransformer(new StreamSource(new StringReader(styleSheet)))
    }
    return transform
}

static processWorksheet(XSSFSheet sheet, Properties prop, writeFile) {
    TransformerFactory transFact = TransformerFactory.newInstance()
    //ClassLoader classLoader = this.getClass().getClassLoader()

    Transformer infoTransform = loadTransformer(transFact, prop, "info-transform")
    Transformer bodyTransform = loadTransformer(transFact, prop, "body-transform")
    Transformer headerTransform = loadTransformer(transFact, prop, "header-transform")
    Transformer resultTransform = loadTransformer(transFact, prop, "result-transform")
    MessageFactory messageFactory = MessageFactory.newInstance()

    def processedLines=0
    boolean isDebug = prop.debug == "true"
    boolean isVerbose = prop.verbose == "true"
    String debugFileName = prop."debug-file"
    PrintStream debugOut
    if (debugFileName != null) {
        debugOut = new PrintStream(new File(debugFileName))
    } else {
        debugOut = System.out
    }



    String ep
    String env = prop.environment
    debugOut.println("env=" + env)
    ep = prop."endpoint-$env" ?: prop.endpoint

    myAssert(ep != null, "Config: environment and/or endpoint must be defined empty")
    debugOut.println("Endpoint: " + ep)

    DocumentBuilderFactory docFactory =
            DocumentBuilderFactory.newInstance()
    DocumentBuilder docBuilder = docFactory.newDocumentBuilder()


    SOAPConnectionFactory soapConnectionFactory =
            SOAPConnectionFactory.newInstance()
    SOAPConnection soapConnection = soapConnectionFactory.createConnection()


    int maxProcess=prop."max-process"?.toInteger()?:1e6
    int maxErrors=prop."max-errors"?.toInteger()?:10
    int commitEvery = prop."commit-every"?.toInteger()?:100

    debugOut.println "Max lines to process: $maxProcess, max errors: $maxErrors, write every: $commitEvery" 
    for (int i = 0; i <= sheet.getLastRowNum() && maxProcess>0 && maxErrors>0; i++) {
        XSSFRow r = sheet.getRow(i)

        prop.setProperty("rownum", Integer.toString(i))
        Document doc = buildRowDoc(docBuilder, r, prop)
        StringWriter inf = new StringWriter()

        infoTransform.transform(new DOMSource(doc), new StreamResult(inf))
        String infoStr = inf.toString()

        if (!"".equals(infoStr)) {
            maxProcess--;
            System.out.print("Row " + (i + 1) + ": " + infoStr + ": ")

            debugOut.println(i + 1 + ": " + infoStr)



            if (isDebug) {
                printXMLDocument(doc, debugOut)
                DOMResult dr = new DOMResult()

                bodyTransform.transform(new DOMSource(doc), dr)
                printXMLDocument((Document) dr.getNode(), debugOut)
            }

            SOAPMessage soapMessage = messageFactory.createMessage()
            SOAPPart soapPart = soapMessage.getSOAPPart()
            SOAPEnvelope envelope = soapPart.getEnvelope()


            SOAPHeader soapHdr = envelope.getHeader()
            SOAPBody soapBody = envelope.getBody()


            bodyTransform.transform(new DOMSource(doc),
                    new DOMResult(soapBody)
            )

            headerTransform.transform(new DOMSource(doc),
                    new DOMResult(soapHdr)
            )

            if (isDebug && isVerbose) {
                printSOAPXML(soapMessage, debugOut)
            }

            soapMessage.saveChanges()
            SOAPMessage soapResponse
            try {
                soapResponse = soapConnection.call(soapMessage, ep)
                if (isDebug) {
                    printSOAPXML(soapResponse, debugOut)
                }



                DOMResult resDom = new DOMResult()
                resultTransform.transform(soapResponse.getSOAPPart().getContent(),
                        resDom
                )

                Document res = (Document) resDom.getNode()
                if (isDebug) {
                    printXMLDocument(res, debugOut)
                }

                if(processedLines++ > commitEvery) {
                    writeFile()
                    processedLines = 0
                }

                if (res != null && res.getDocumentElement() != null) {
                    NodeList cols =
                            res.getDocumentElement().getChildNodes()
                    for (int k = 0; k < cols.getLength(); k++) {
                        Node cx = cols.item(k)
                        int colnr =
                                Integer.parseInt(cx.getAttributes().getNamedItem("col").getTextContent())
                        setOrCreateCell(r, colnr, cx.getTextContent())
                    }
                } else {
                    maxErrors--;
                    setOrCreateCell(r, 0, "ERROR")
                    setOrCreateCell(r, 1, "Result after transformation is empty, check logs")
                }
            }
            catch (Exception e) {
                maxErrors--;
                e.printStackTrace(System.out)
                println(e.getLocalizedMessage())
                setOrCreateCell(r, 0, "ERROR")
                setOrCreateCell(r, 1, e.getLocalizedMessage())
            }
        }
    }
}


static loadOverrideProperties(XSSFSheet sheet, Properties prop) {
    if (sheet == null) {
        return
    }


    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
        XSSFRow r = sheet.getRow(i)
        if (r != null) {
            String k = getCell(r, 0)
            String v = getCell(r, 1)
            if (k != null && k != "" && v != null) {
                prop[k] = v
            }
        }
    }
}


static void main(String[] args) {

    println("Spreadsheet2WS by jramb")
    println("https://github.com/jramb/spreadsheet2ws")

    String filename
    Properties prop = new Properties()
    ClassLoader classLoader = this.getClass().getClassLoader()
    InputStream config =
            classLoader.getResourceAsStream("config.xml") // config.properties?

    //myAssert(config != null, "Could not find config.xml in path, I need that!");
    if (config != null) {
        prop.loadFromXML(config)
        System.err.println("Info: config.xml loaded from path")
    } else {
        System.err.println("Info: config.xml not loaded from path")
    }

    //prop.setProperty("now", new Date().format("yyyy-MM-dd"))
    prop.now = new Date().format("yyyy-MM-dd")
    prop.uuid = UUID.randomUUID() as String

    if (args.size() == 0) {
      println '''** You need to specify the spreadsheet to load as the first parameter **

Usage:
  run filename.xlsx [param=value]*

'''
      System.exit(-1)
    }

    int i = 0
    if (args.length > 0) {
        File f = new File(args[i])
        if (f.exists()) {
            prop.excelfile = f.getName()
            i++
        }
    }

    filename = prop.excelfile
    println("Starting spreadsheet processing: " + filename)
    File xlsx = new File(filename)
    File backup = new File(xls.absoulteFile+".bak")
    myAssert(xlsx.canRead(), "File not found: " + xlsx.getName())
    myAssert(xlsx.canWrite(), "File not writeable: " + xlsx.getName())
    if(updateFile) {
        if (backup.exists()) {
          assert backup.delete(), "Could not remove old backup"
        }
    }

    XSSFWorkbook x = new XSSFWorkbook(new FileInputStream(xlsx))

    loadOverrideProperties(x.getSheetAt(1), prop)

    // apply args, override all
    for ( /* continue using i */ ; i < args.length; i++) {
        String[] v = args[i].split("=")
        if (v.length == 2) {
            prop.setProperty(v[0], v[1])
        }
    }

    // make sure we can write the file by... writing it
    boolean updateFile = prop."update-file" != "false"

    def writeFile = {
        if(updateFile) {
            print "Writing file ${xlsx.name}"
            File tmp = new File(xlsx.absoluteFile+".tmp")
            //x.write(new FileOutputStream(xlsx))
            if (!backup.exists()) {
                assert xlsx.renameTo(backup), "Could not backup file!"
            }
            x.write(new FileOutputStream(tmp))
            if(xlsx.exists()) {
                assert xlsx.delete(), "Could not delete ${xlsx.name} before replacing"
            }
            assert tmp.renameTo(xlsx), "Could not restore written file, see backup ${backup.name}"
            print "."
        }
    }

    try {
        writeFile();
        //x.write(new FileOutputStream(xlsx))

    }

    catch (IOException e) {
        println("Can not write to " + filename)
        println("Maybe you need to close the file in Excel?")
        System.exit(1)
    }
    try {
        processWorksheet(x.getSheetAt(0), prop, writeFile )
    }
    finally {
        if (updateFile) {
            println("Updating " + xlsx.getName())
            writeFile(); //x.write(new FileOutputStream(xlsx))
        }
    }

    println "Done."
}




