package com.viacerta.controller;

import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.eclipse.microprofile.openapi.annotations.media.Content;
import org.eclipse.microprofile.openapi.annotations.media.Schema;
import org.eclipse.microprofile.openapi.annotations.responses.APIResponse;
import org.eclipse.microprofile.openapi.annotations.tags.Tag;
import org.jboss.resteasy.annotations.providers.multipart.MultipartForm;

import com.viacerta.DTO.FileUploadForm;
import com.viacerta.service.FreemarkerProcessorService;
import com.viacerta.service.LibreOfficeConverterService;

@Path("/generate")
@Tag(name = "Document Generation", description = "API para geração de documentos com placeholders e regras de exibição de parágrafos.")
public class DocumentController {

    @POST
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces(MediaType.APPLICATION_OCTET_STREAM)
        @APIResponse(
        responseCode = "200",
        description = "Arquivo PDF gerado",
        content = @Content(mediaType = "application/octet-stream", schema = @Schema(implementation = File.class))
    )
    @APIResponse(responseCode = "500", description = "Erro interno do servidor")
    public Response generateDocument(@MultipartForm FileUploadForm form) {
        try {
            if (form.getDocFile() == null || form.getFieldsJson() == null) {
                return Response.status(Response.Status.BAD_REQUEST)
                               .entity("Dados de entrada incompletos")
                               .build();
            }
            InputStream docxInputStream = form.getDocFile();
            // 1. Carregar o documento Word
            XWPFDocument document = new XWPFDocument(docxInputStream);

            // 2. Processar placeholders com Freemarker
            FreemarkerProcessorService freemarkerProcessor = new FreemarkerProcessorService();
            freemarkerProcessor.processTemplate(document, form.getFieldsJson(), form.getRulesJson());

            // 3. Converter para PDF usando LibreOffice (modo headless)
            File pdfFile = LibreOfficeConverterService.convertToPDF(document);

            if (pdfFile == null || !pdfFile.exists()) {
                throw new IOException("Falha ao gerar o arquivo PDF");
            }

            return Response.ok(pdfFile)
                .header("Content-Disposition", "attachment; filename=\"generated_document.pdf\"")
                .build();
        } catch (IOException e) {
            e.printStackTrace(); // Log detalhado do erro
            return Response.status(Response.Status.INTERNAL_SERVER_ERROR)
                .entity("Erro ao processar o arquivo: " + e.getMessage())
                .build();
        } catch (Exception e) {
            e.printStackTrace(); // Log detalhado do erro
            return Response.status(Response.Status.INTERNAL_SERVER_ERROR)
                .entity("Erro inesperado: " + e.getMessage())
                .build();
        }
    }
}
