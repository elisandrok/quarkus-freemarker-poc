package com.viacerta;

import java.io.InputStream;
import java.util.Map;

import org.jboss.resteasy.reactive.PartType;

import jakarta.ws.rs.FormParam;
import jakarta.ws.rs.core.MediaType;

public class WordDataForm {
    @FormParam("file")
    @PartType(MediaType.APPLICATION_OCTET_STREAM)
    private InputStream file;

    @FormParam("variables")
    @PartType(MediaType.APPLICATION_JSON)
    private Map<String, Object> variables;

    public InputStream getFile() {
        return file;
    }

    public Map<String, Object> getVariables() {
        return variables;
    }
    
}
