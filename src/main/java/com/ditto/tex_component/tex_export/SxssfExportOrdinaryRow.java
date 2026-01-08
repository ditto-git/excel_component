package com.ditto.tex_component.tex_export;

import com.ditto.tex_component.tex_exception.TexException;
import com.ditto.tex_component.tex_util.TexThreadLocal;
import com.ditto.tex_component.tex_util.oss.OSSInputOperate;
import com.ditto.tex_component.tex_util.oss.OSSUtil;
import com.ditto.tex_component.tex_util.request.ExportFileResponseUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;

import static com.ditto.tex_component.tex_exception.TexExceptionEnum.FILE_EXPORT_ERROR;


@Component("SxssfExportOrdinary2")
@Slf4j
public class SxssfExportOrdinaryRow implements SxssfExportOrdinary{



    @Override
    public void export( HttpServletResponse response, GoExport goExport) {
        ExportFileResponseUtil responseUtil = new ExportFileResponseUtil(response, TexThreadLocal.getExTemplate().getFileName(), "xlsx");
        OSSUtil.downloadOSSInput(TexThreadLocal.getExTemplate().getTemplateUrl(), new OSSInputOperate() {
            @Override
            public void closeBefore(InputStream inputStream) throws Exception {
                SxssfExport exportColum = SxssfExportFactory.create(inputStream, TexThreadLocal.getExTemplate().getTemplateType());
                goExport.exportData(exportColum);
                TexThreadLocal.clear();
                try (ServletOutputStream outputStream = responseUtil.getOutputStream()) {
                    exportColum.getWorkbook().write(outputStream);
                    exportColum.getWorkbook().dispose();
                } catch (Exception e) {
                    log.error(e.getMessage(), e);
                    throw new TexException(FILE_EXPORT_ERROR);
                }
            }

            @Override
            public void closeAfter() throws Exception {

            }
        });




    }

    @Override
    public void exportLocalFile(HttpServletResponse response, GoExport goExport) {

    }


}
