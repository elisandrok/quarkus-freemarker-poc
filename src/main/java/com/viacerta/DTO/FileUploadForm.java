package com.viacerta.DTO;

import java.io.InputStream;

import jakarta.ws.rs.FormParam;
import jakarta.ws.rs.core.MediaType;
import java.util.Map;

import org.jboss.resteasy.annotations.providers.multipart.PartType;


public class FileUploadForm  {
    @FormParam("file")
    @PartType(MediaType.APPLICATION_OCTET_STREAM)
    public InputStream docFile;

    @FormParam("fileSubReport")
    @PartType(MediaType.APPLICATION_OCTET_STREAM)
    public InputStream docFileSubReport;

    @FormParam("fieldsJson")
    @PartType(MediaType.APPLICATION_JSON)
    public Map<String, Object> fieldsJson;

    @FormParam("rulesJson")
    @PartType(MediaType.APPLICATION_JSON)
    public Map<String, Boolean> rulesJson;

    public InputStream getDocFile() {
        return docFile;
    }

    public InputStream getDocFileSubReport() {
        return docFileSubReport;
    }

    public Map<String, Object> getFieldsJson() {
        return fieldsJson;
    }

    public Map<String, Boolean> getRulesJson() {
        return rulesJson;
    }
}
