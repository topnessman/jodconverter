//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.task.ErrorCodeIOException;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;
import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeTask;
import org.artofsolving.jodconverter.streams.OOoInputStream;
import org.artofsolving.jodconverter.streams.OOoOutputStream;

import java.util.HashMap;
import java.util.Map;

import static org.artofsolving.jodconverter.office.OfficeUtils.SERVICE_DESKTOP;
import static org.artofsolving.jodconverter.office.OfficeUtils.URL_STREAM;
import static org.artofsolving.jodconverter.office.OfficeUtils.cast;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUnoProperties;

public class StreamConversionTask implements OfficeTask {

    private final OOoInputStream ooInputStream;
    private final OOoOutputStream ooOutputStream;
    private final String filterName;
    public StreamConversionTask(OOoInputStream ooInputStream, OOoOutputStream ooOutputStream, String filterName) {
        this.ooInputStream = ooInputStream;
        this.ooOutputStream = ooOutputStream;
        this.filterName = filterName;
    }

    protected Map<String,?> getLoadProperties(OOoInputStream ooInputStream) {
        Map<String,Object> loadProperties = new HashMap<String,Object>();
        loadProperties.put("InputStream", ooInputStream);
        loadProperties.put("Hidden", new Boolean(true));
        return loadProperties;
    };

    protected Map<String,?> getStoreProperties(OOoOutputStream ooOutputStream, XComponent document) {
        Map<String,Object> storeProperties = new HashMap<String,Object>();
        storeProperties.put("OutputStream", ooOutputStream);
        storeProperties.put("FilterName", filterName);
        return storeProperties;
    };

    public void execute(OfficeContext context) throws OfficeException {
        XComponent document = null;
        try {
            document = loadDocument(context, ooInputStream);
            modifyDocument(document);
            storeDocument(document, ooOutputStream);
        } catch (OfficeException officeException) {
            throw officeException;
        } catch (Exception exception) {
            throw new OfficeException("conversion failed", exception);
        } finally {
            if (document != null) {
                XCloseable closeable = cast(XCloseable.class, document);
                if (closeable != null) {
                    try {
                        closeable.close(true);
                    } catch (CloseVetoException closeVetoException) {
                        // whoever raised the veto should close the document
                    }
                } else {
                    document.dispose();
                }
            }
        }
    }

    private XComponent loadDocument(OfficeContext context, OOoInputStream ooInputStream) throws OfficeException {
        if (ooInputStream == null) {
            throw new OfficeException("input bytes are null");
        }
        XComponentLoader loader = cast(XComponentLoader.class, context.getService(SERVICE_DESKTOP));
        Map<String,?> loadProperties = getLoadProperties(ooInputStream);
        XComponent document = null;
        try {
            document = loader.loadComponentFromURL(URL_STREAM, "_blank", 0, toUnoProperties(loadProperties));
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new OfficeException("could not load XComponent from : "  + ooInputStream, illegalArgumentException);
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not load XComponent from : "  + ooInputStream + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not load XComponent from : "  + ooInputStream, ioException);
        }
        if (document == null) {
            throw new OfficeException("could not load XComponent from : "  + ooInputStream);
        }
        return document;
    }

    /**
     * Override to modify the document after it has been loaded and before it gets
     * saved in the new format.
     * <p>
     * Does nothing by default.
     * 
     * @param document
     * @throws OfficeException
     */
    protected void modifyDocument(XComponent document) throws OfficeException {
    	// noop
    }

    private void storeDocument(XComponent document, OOoOutputStream ooOutputStream) throws OfficeException {
        Map<String,?> storeProperties = getStoreProperties(ooOutputStream, document);
        if (storeProperties == null) {
            throw new OfficeException("unsupported conversion");
        }
        try {
            cast(XStorable.class, document).storeToURL(URL_STREAM, toUnoProperties(storeProperties));
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not store document to : " + ooOutputStream + "; errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not store document to : " + ooOutputStream, ioException);
        }
    }

}
