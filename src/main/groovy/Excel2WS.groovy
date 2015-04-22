//package xxcust.spread2ws
/*
 * 2014 by J Ramb, Navigate Consulting
 */

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


def myAssert(boolean cond, String message) {
  if (!cond) {
    println message
      System.exit(1)
  }
}


def streamDOMSource(Source ds, StreamResult sr) {

  Transformer transformer = 
    TransformerFactory.newInstance().newTransformer()
  transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes")
  transformer.setOutputProperty(OutputKeys.METHOD, "xml")
  transformer.setOutputProperty(OutputKeys.INDENT, "yes")
  transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8")
  transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4")

  transformer.transform(ds, sr)
}

def printXMLDocument(Document doc, OutputStream out)
{
  streamDOMSource(new DOMSource(doc), 
      new StreamResult(new OutputStreamWriter(out, "UTF-8")))
}

def printSOAPXML(SOAPMessage soapResponse, PrintStream out) {
  streamDOMSource(soapResponse.getSOAPPart().getContent(), new StreamResult(out))
}


String xmlToString(Document doc) {
  StringWriter sw = new StringWriter()
  streamDOMSource(new DOMSource(doc), new StreamResult(sw))
  return sw.toString()
}



// Oracles NVL
String nvl(String val, String ifnull) {
  val == null || "".equals(val.trim())?ifnull:val
}



XSSFCell setOrCreateCell(XSSFRow r, int col, String val) {
  XSSFCell c = r.getCell(col)
    if (c == null) {
      c = r.createCell(col)
    }
  c.setCellValue(val)
    if (col == 0) {
      println(val)
    }
  return c
}


String getCell(XSSFRow r, int col) {
  XSSFCell c = r.getCell(col)
  if (c == null) { 
    return null
  }
  switch (c.getCellType()) {
    case Cell.CELL_TYPE_NUMERIC:
      return new BigDecimal(c.getNumericCellValue()).toString()
    case Cell.CELL_TYPE_BOOLEAN:
      return c.getBooleanCellValue()?"true":"false"
    case Cell.CELL_TYPE_STRING:
      return c.getStringCellValue()
    default:
            return c.getStringCellValue()
  }
}


Document buildRowDoc(DocumentBuilder docBuilder, XSSFRow r, Properties prop) {
  Document doc = docBuilder.newDocument()
  Element rootElement = doc.createElement("root")
  doc.appendChild(rootElement)
  //rootElement.appendChild(doc.importNode(configXML,true)); // IMPORT the alien XML... DOES NOT WORK??
  // ok, doing it by hand {{{

    Element properties = doc.createElement("properties")
    rootElement.appendChild(properties)
    for (Enumeration e = prop.propertyNames() ; e.hasMoreElements()
        ; ){

      String s = (String)e.nextElement()
        Element ent = doc.createElement("entry")
        ent.setAttribute("key", s)
        ent.setTextContent(prop.getProperty(s))
        properties.appendChild(ent)
    }
    // ok, doing it by hand }}}

  Element rowElement = doc.createElement("row")
  rootElement.appendChild(rowElement)

  for (int l = 0 ; l <= r.getLastCellNum() ; l++){
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


void processWorksheet(XSSFSheet sheet, Properties prop) {
  TransformerFactory transFact = TransformerFactory.newInstance()
  ClassLoader classLoader = Excel2WS.getClassLoader()

  Transformer infoTransform = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("info-transform"))))
  Transformer bodyTransform = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("body-transform"))))
  Transformer headerTransform = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("header-transform"))))
  Transformer resultTransform = transFact.newTransformer(new StreamSource(classLoader.getResourceAsStream(prop.getProperty("result-transform"))))
  MessageFactory messageFactory = MessageFactory.newInstance()

  boolean isDebug = "true".equals(prop.getProperty("debug"))
  String debugFileName = prop.getProperty("debug-file")
  PrintStream debugOut
  if (debugFileName != null) {
    debugOut = new PrintStream(new File(debugFileName))
  } else {
    debugOut = System.out
  }



  String ep
  String env = prop.getProperty("environment")
  debugOut.println("env=" + env)
  ep = nvl(prop.getProperty("endpoint-" + env), prop.getProperty("endpoint"))

  myAssert(ep != null, "Config: environment and/or endpoint must be defined empty")
  debugOut.println("Endpoint: " + ep)

  DocumentBuilderFactory docFactory = 
  DocumentBuilderFactory.newInstance()
  DocumentBuilder docBuilder = docFactory.newDocumentBuilder()


  SOAPConnectionFactory soapConnectionFactory = 
  SOAPConnectionFactory.newInstance()
  SOAPConnection soapConnection = soapConnectionFactory.createConnection()


  for (int i = 0 ; i <= sheet.getLastRowNum() ; i++){
    XSSFRow r = sheet.getRow(i)

    prop.setProperty("rownum", Integer.toString(i))
    Document doc = buildRowDoc(docBuilder, r, prop)
    StringWriter inf = new StringWriter()

    infoTransform.transform(new DOMSource(doc), new StreamResult(inf))
    String infoStr = inf.toString()

    if (!"".equals(infoStr)) {
      System.out.print("Row " + i + 1 + ": " + infoStr + ": ")

        debugOut.println(i + 1 + ": " + infoStr)



          if (isDebug) {

              printXMLDocument(doc, debugOut)
                DOMResult dr = new DOMResult()

                bodyTransform.transform(new DOMSource(doc), dr)
                printXMLDocument((Document)dr.getNode(), debugOut)
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

            if (isDebug) {
              printSOAPXML(soapMessage, debugOut)
            }




          soapMessage.saveChanges()
            SOAPMessage soapResponse = null
            try {
              soapResponse = soapConnection.call(soapMessage, ep)
                if (isDebug) {
                  printSOAPXML(soapResponse, debugOut)
                }



              DOMResult resDom = new DOMResult()
                resultTransform.transform(soapResponse.getSOAPPart().getContent(), 
                    resDom
                    )

                Document res = (Document)resDom.getNode()
                if (isDebug) {
                  printXMLDocument(res, debugOut)
                }





              if (res != null && res.getDocumentElement() != null) {
                NodeList cols = 
                  res.getDocumentElement().getChildNodes()
                  for (int k = 0 ; k < cols.getLength() ; k++){
                    Node cx = cols.item(k)
                      int colnr = 
                      Integer.parseInt(cx.getAttributes().getNamedItem("col").getTextContent())
                      setOrCreateCell(r, colnr, cx.getTextContent())
                  }
              } else {
                setOrCreateCell(r, 0, "ERROR")
                  setOrCreateCell(r, 1, "Result after transformation is empty, check logs")
              }
            }
          catch (Exception e) {
            e.printStackTrace(System.out)
              println(e.getLocalizedMessage())
              setOrCreateCell(r, 0, "ERROR")
              setOrCreateCell(r, 1, e.getLocalizedMessage())
          }
        }
    }
}




def loadOverrideProperties(XSSFSheet sheet, Properties prop) {
  if (sheet == null) {
    return
  }


  for (int i = 0 ; i <= sheet.getLastRowNum() ; i++){
    XSSFRow r = sheet.getRow(i)
    if (r != null) {
      String k = getCell(r, 0)
      String v = getCell(r, 1)
      if (!"".equals(k)) {
        prop.setProperty(k, v)
      }
    }
  }
}



static void main(String[] args) {

  println("excelwsload by jramb")
  String filename
  Properties prop = new Properties()
  ClassLoader classLoader = Excel2WS.getClassLoader()
  InputStream config = 
    classLoader.getResourceAsStream("config.xml") // config.properties?


    //myAssert(config != null, "Could not find config.xml in path, I need that!");
    if (config != null) {
      prop.loadFromXML(config)
    } else {
      System.err.println("config.xml not loaded from path")
    }





  prop.setProperty("now", new Date().format("yyyy-MM-dd"))

  prop.setProperty("uuid", UUID.randomUUID())

  int i = 0
  if (args.length > 0) {
    File f = new File(i[])
      if (f.exists()) {
        prop.setProperty("excelfile", f.getName())
        i++
      }
  }







  filename = prop.getProperty("excelfile")
  println("Starting spreadsheet processing: " + filename)
  File xlsx = new File(filename)
  myAssert(xlsx.canRead(), "File not found: " + xlsx.getName())
  myAssert(xlsx.canWrite(), "File not writeable: " + xlsx.getName())

  XSSFWorkbook x = new XSSFWorkbook(new FileInputStream(xlsx))

  loadOverrideProperties(x.getSheetAt(1), prop)


  // apply args, override all
  for ( /* continue using i */; i < args.length ; i++){
    String[] v = i[].split("=")
      if (v.length == 2) {
        prop.setProperty(v[0], v[1])
      }
  }


  // make sure we can write the file by... writing it
  boolean updateFile = !"false".equals(prop.getProperty("update-file"))

  try {
    if (updateFile) {
      x.write(new FileOutputStream(xlsx))
    }
  }

  catch (IOException e) {
    println("Can not write to " + filename)
      println("Maybe you need to close the file in Excel?")
      System.exit(1)
  }
  try {
    processWorksheet(x.getSheetAt(0), prop)
  }
  finally {
    if (updateFile) {
      println("Updating " + xlsx.getName())
        x.write(new FileOutputStream(xlsx))
    }
  }

  println "Done."
}



