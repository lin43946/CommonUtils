

import com.aliyun.oss.ClientException;
import com.aliyun.oss.OSSClient;
import com.aliyun.oss.OSSException;
import com.aliyun.oss.model.CannedAccessControlList;
import com.aliyun.oss.model.CreateBucketRequest;
import com.bgyfw.comprehensive.config.OssInfo;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.util.Date;

@Component("ossUtils")
@Slf4j
public class OSSUtils {
	


	public OSSClient createCilent(){
        String endpoint =OssInfo.END_POINT;
        String accessKeyId =OssInfo.ACCESS_KEY_ID;
        String accessKeySecret =OssInfo.ACCESS_KEY_SECRET;
        OSSClient client = new OSSClient(endpoint, accessKeyId, accessKeySecret);
        return client;
    }
	
	/**
	 * 单个文件上传，大小自行判断，默认不超过5G。
	 * @param fileName 文件名
	 * @param inputStream 文件流对象
	 * @return 返回下载url
	 */
    public String uploadFile(String fileName, InputStream inputStream){
       OSSClient ossClient = createCilent();
       String url = "";
       try {
       		//容器不存在，就创建
		   	if(!ossClient.doesBucketExist(OssInfo.BUCKET_NAME)){
				ossClient.createBucket(OssInfo.BUCKET_NAME);
				CreateBucketRequest createBucketRequest = new CreateBucketRequest(OssInfo.BUCKET_NAME);
				createBucketRequest.setCannedACL(CannedAccessControlList.PublicRead);
				ossClient.createBucket(createBucketRequest);
			}
    	    fileName = OssInfo.FILE_DIR + "/"+ fileName;
    	    boolean doesObjectExist = ossClient.doesObjectExist(OssInfo.BUCKET_NAME,fileName);
    	    if(!doesObjectExist) {
    	    	ossClient.putObject(OssInfo.BUCKET_NAME, fileName, inputStream);
                url=ossClient.generatePresignedUrl(OssInfo.BUCKET_NAME,fileName,new Date(System.currentTimeMillis() + 3600L * 1000 * 24 * 365 * 50)).toString();
    	    }
            // 上传文件流
            
        } catch (OSSException oe) {
        	log.error("{}=Caught an OSSException, which means your request made it to OSS,"
        			+ "but was rejected with an error response for some reason.", oe.getMessage(), oe);
        	log.error("{}=Error Message: "+ oe.getErrorCode());
        	log.error("{}=Error Message: " + oe.getErrorCode());
        	log.error("{}=Error Code:       " + oe.getErrorCode());
        	log.error("{}=Request ID:      " + oe.getRequestId());
        	log.error("{}=Host ID:           " + oe.getHostId());
        } catch (ClientException ce) {
        	log.error("{}=Caught an ClientException, which means the client encountered "
        			+ "a serious internal problem while trying to communicate with OSS, "
        			+ "such as not being able to access the network.");
        	log.error("{}=Error Message: " + ce.getMessage());
        } finally {
            if (ossClient != null) {
            	ossClient.shutdown();
            }
        }
        return StringUtils.isEmpty(url)?"文件已存在":url;
    }



	public String uploadFileExpired(String fileName, InputStream inputStream,Long expiretime){
		OSSClient ossClient = createCilent();
		String url = "";
		try {
			fileName = OssInfo.FILE_DIR + "/"+ fileName;
			boolean doesObjectExist = ossClient.doesObjectExist(OssInfo.BUCKET_NAME,fileName);
			if(!doesObjectExist) {
				ossClient.putObject(OssInfo.BUCKET_NAME, fileName, inputStream);
				url=ossClient.generatePresignedUrl(OssInfo.BUCKET_NAME,fileName,new Date(System.currentTimeMillis() + expiretime)).toString();
			}
			// 上传文件流

		} catch (OSSException oe) {
			log.error("{}=Caught an OSSException, which means your request made it to OSS,"
					+ "but was rejected with an error response for some reason.", oe.getMessage(), oe);
			log.error("{}=Error Message: "+ oe.getErrorCode());
			log.error("{}=Error Message: " + oe.getErrorCode());
			log.error("{}=Error Code:       " + oe.getErrorCode());
			log.error("{}=Request ID:      " + oe.getRequestId());
			log.error("{}=Host ID:           " + oe.getHostId());
		} catch (ClientException ce) {
			log.error("{}=Caught an ClientException, which means the client encountered "
					+ "a serious internal problem while trying to communicate with OSS, "
					+ "such as not being able to access the network.");
			log.error("{}=Error Message: " + ce.getMessage());
		} finally {
			if (ossClient != null) {
				ossClient.shutdown();
			}
		}
		return StringUtils.isEmpty(url)?"文件已存在":url;
	}
    
    /**
     * 删除文件
     * @param fileName
     */
    public void delFile(String fileName) {
    	OSSClient ossClient = createCilent();
    	try {
        	ossClient.deleteObject(OssInfo.BUCKET_NAME, OssInfo.FILE_DIR+"/"+fileName);
        	ossClient.shutdown();
		} catch (OSSException e) {
			log.error("{}=Error Message: " + e.getMessage());
			log.error("{}=Error Message: 文件来自旧服务器。");
		} finally {
            if (ossClient != null) {
            	ossClient.shutdown();
            }
        }
    	
    	
    }

}
