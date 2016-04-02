package itstack.demo.upload.web;

import itstack.demo.upload.common.GsonUtils;
import itstack.demo.upload.common.PoiUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@Controller("upload")
@RequestMapping("/upload/")
public class UploadController {

    private Logger logger = LoggerFactory.getLogger(UploadController.class);

    @ResponseBody
    @RequestMapping(value = "uploadFile", method = RequestMethod.POST, produces = "text/html;charset=utf-8")
    public boolean uploadFile(@RequestParam("file") CommonsMultipartFile file, HttpServletRequest req, HttpServletResponse res) throws IOException {
        List<String[]> list = PoiUtils.readRecordsInputStream(file.getFileItem().getInputStream(), 0, true, 1);
        logger.info("读取结果：" + GsonUtils.toJson(list));
        /**
         * 如果需要上传到指定目录
         * //设置文件保存的本地路径
           String filePath = req.getSession().getServletContext().getRealPath("/uploadFiles/");
           String fileName = pic.getOriginalFilename();
           String fileType = fileName.split("[.]")[1];
           //为了避免文件名重复，在文件名前加UUID
           String uuid = UUID.randomUUID().toString().replace("-", "");
           String uuidFileName = uuid + fileName;
           File f = new File(filePath + "/" + uuid + "." + fileType);
           //将文件保存到服务器
           FileUtil.upFile(pic.getInputStream(), uuidFileName, filePath);
         */
        return true;
    }

}
