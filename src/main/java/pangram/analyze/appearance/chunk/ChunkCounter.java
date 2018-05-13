package pangram.analyze;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;

public class ChunkCounter extends DefaultHandler {
    private static int countChunk = 0;

    static int getCountChunk(){
        return countChunk;
    }

    static void build(InputStream in) throws IOException {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        try {
            SAXParser parser = factory.newSAXParser();
            parser.parse(in, new ChunkCounter());
        }catch (SAXException | ParserConfigurationException e){
            throw new RuntimeException(e);
        }
    }

    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) {
        if(qName.equals("Chunk")) {
            countChunk++;
        }
    }
}
