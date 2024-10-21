package com.viacerta.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.eclipse.microprofile.openapi.annotations.tags.Tag;

import com.viacerta.DTO.FileUploadForm;
import com.viacerta.service.LibreOfficeConverterService;
import com.viacerta.service.ProcessDocumentWithFreemarkerService;

import jakarta.inject.Inject;
import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;

import org.jboss.resteasy.annotations.providers.multipart.MultipartForm;

@Path("/generateV2")
@Tag(name = "Document Generation V2", description = "API para geração de relatórios a partir de templates word exportando em pdf.")
public class DocumentController {

    private ProcessDocumentWithFreemarkerService serviceFreemarker;
    private LibreOfficeConverterService libreOfficeConverterService;

    @Inject
    public DocumentController(ProcessDocumentWithFreemarkerService serviceFreemarker,

            LibreOfficeConverterService libreOfficeConverterService) {
        this.serviceFreemarker = serviceFreemarker;
        this.libreOfficeConverterService = libreOfficeConverterService;
    }

    @POST
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces(MediaType.APPLICATION_OCTET_STREAM)
    public Response generateReport(@MultipartForm FileUploadForm form) {
        try {
            if (form.getDocFile() == null || form.getFieldsJson() == null) {
                return Response.status(Response.Status.BAD_REQUEST)
                        .entity("Dados de entrada incompletos")
                        .build();
            }
            XWPFDocument document = new XWPFDocument(form.getDocFile());

            serviceFreemarker.processDocumentWithFreemarker(document, form.getFieldsJson());

            File pdfFile = libreOfficeConverterService.convertToPDF(document);
            File subReportPdfFile = new File("subReport.pdf");

            if (form.getDocFileSubReport() == null || form.getDocFileSubReport().available() == 0) {
                Response.ResponseBuilder response = Response.ok((Object) new FileInputStream(pdfFile));
                response.header("Content-Disposition", "attachment; filename=\"relatorio.pdf\"");
                response.type("application/pdf");

                pdfFile.delete();
                return response.build();
            }
            XWPFDocument subReportDocument = new XWPFDocument(form.getDocFileSubReport());
            //Processar o subrelatório
            serviceFreemarker.processDocumentWithFreemarker(subReportDocument, form.getFieldsJson());
            subReportPdfFile = libreOfficeConverterService.convertToPDF(subReportDocument);
            byte[] mergedPdf = libreOfficeConverterService.unifyDocuments(pdfFile,subReportPdfFile);

            if (mergedPdf == null || mergedPdf.length == 0) {
                throw new IOException("Falha ao gerar o arquivo PDF");
            }

            Response.ResponseBuilder response = Response.ok((Object) new FileInputStream(pdfFile));
            response.header("Content-Disposition", "attachment; filename=\"relatorio.pdf\"");
            response.type("application/pdf");

            pdfFile.delete();
            subReportPdfFile.delete();

            return response.build();
        } catch (Exception e) {
            e.printStackTrace();
            return Response.status(Response.Status.INTERNAL_SERVER_ERROR).entity("Erro ao gerar o relatório").build();
        }
    }
}