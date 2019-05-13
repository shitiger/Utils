package httpPool;


import org.apache.commons.lang.StringUtils;
import org.apache.http.HeaderElement;
import org.apache.http.HeaderElementIterator;
import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.methods.*;
import org.apache.http.client.protocol.HttpClientContext;
import org.apache.http.conn.ConnectionKeepAliveStrategy;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.impl.conn.PoolingHttpClientConnectionManager;
import org.apache.http.message.BasicHeaderElementIterator;
import org.apache.http.protocol.HTTP;
import org.apache.http.protocol.HttpContext;
import org.apache.http.util.EntityUtils;

import java.text.ParseException;

/**
 * @author stone tiger
 * @Description: ${todo}
 * @date 2019/3/25
 */
public class HttpClientPoolUtil {


    public static PoolingHttpClientConnectionManager cm = null;

    public static CloseableHttpClient httpClient = null;

    /**
     * 默认content 类型
     */
    private static final String DEFAULT_CONTENT_TYPE = "application/json";

    /**
     * 默认请求超时时间30s
     */
    private static final int DEFAUL_TTIME_OUT = 3000;

    private static final int count = 10;

    private static final int totalCount = 20;

    private static final int Http_Default_Keep_Time = 1000;

    /**
     * 初始化连接池
     */
    public static synchronized void initPools() {
        if (httpClient == null) {
            cm = new PoolingHttpClientConnectionManager();
            cm.setDefaultMaxPerRoute(count);
            cm.setMaxTotal(totalCount);
            httpClient = HttpClients.custom().setKeepAliveStrategy(defaultStrategy).setConnectionManager(cm).build();
        }
    }

    /**
     * Http connection keepAlive 设置
     */
    public static ConnectionKeepAliveStrategy defaultStrategy = new ConnectionKeepAliveStrategy() {
        public long getKeepAliveDuration(HttpResponse response, HttpContext context) {
            HeaderElementIterator it = new BasicHeaderElementIterator(response.headerIterator(HTTP.CONN_KEEP_ALIVE));
            int keepTime = Http_Default_Keep_Time;
            while (it.hasNext()) {
                HeaderElement he = it.nextElement();
                String param = he.getName();
                String value = he.getValue();
                if (value != null && param.equalsIgnoreCase("timeout")) {
                    try {
                        return Long.parseLong(value) * 1000;
                    } catch (Exception e) {
                        e.printStackTrace();
                     //   logger.error("format KeepAlive timeout exception, exception:" + e.toString());
                    }
                }
            }
            return keepTime * 1000;
        }
    };

    public static CloseableHttpClient getHttpClient() {
        return httpClient;
    }

    public static PoolingHttpClientConnectionManager getHttpConnectionManager() {
        return cm;
    }

    /**
     * 执行http post请求 默认采用Content-Type：application/json，Accept：application/json
     *
     * @param uri 请求地址
     * @param data  请求数据
     * @return
     */
    public static String execute(String uri, String data) {
        long startTime = System.currentTimeMillis();
        HttpEntity httpEntity = null;
        HttpEntityEnclosingRequestBase method = null;
        String responseBody = "";
        try {
            if (httpClient == null) {
                initPools();
            }
            method = (HttpEntityEnclosingRequestBase) getRequest(uri, HttpPost.METHOD_NAME, DEFAULT_CONTENT_TYPE, 0);
            method.setEntity(new StringEntity(data));
            HttpContext context = HttpClientContext.create();
            CloseableHttpResponse httpResponse = httpClient.execute(method, context);
            httpEntity = httpResponse.getEntity();
            if (httpEntity != null) {
                responseBody = EntityUtils.toString(httpEntity, "UTF-8");
            }

        } catch (Exception e) {
            if(method != null){
                method.abort();
            }
            e.printStackTrace();
          //  logger.error("execute post request exception, url:" + uri + ", exception:" + e.toString() + ", cost time(ms):"+ (System.currentTimeMillis() - startTime));
        } finally {
            if (httpEntity != null) {
                try {
                    EntityUtils.consumeQuietly(httpEntity);
                } catch (Exception e) {
                    e.printStackTrace();
                   // logger.error(
                      //      "close response exception, url:" + uri + ", exception:" + e.toString() + ", cost time(ms):"
                          //          + (System.currentTimeMillis() - startTime));
                }
            }
        }
        return responseBody;
    }

    /**
     * 创建请求
     *
     * @param uri 请求url
     * @param methodName 请求的方法类型
     * @param contentType contentType类型
     * @param timeout 超时时间
     * @return
     */
    public static HttpRequestBase getRequest(String uri, String methodName, String contentType, int timeout) {
        if (httpClient == null) {
            initPools();
        }
        HttpRequestBase method = null;
        if (timeout <= 0) {
            timeout = DEFAUL_TTIME_OUT;
        }
        RequestConfig requestConfig = RequestConfig.custom().setSocketTimeout(timeout * 1000).setConnectTimeout(timeout * 1000)
                .setConnectionRequestTimeout(timeout * 1000).setExpectContinueEnabled(false).build();

        if (HttpPut.METHOD_NAME.equalsIgnoreCase(methodName)) {
            method = new HttpPut(uri);
        } else if (HttpPost.METHOD_NAME.equalsIgnoreCase(methodName)) {
            method = new HttpPost(uri);
        } else if (HttpGet.METHOD_NAME.equalsIgnoreCase(methodName)) {
            method = new HttpGet(uri);
        } else {
            method = new HttpPost(uri);
        }
        if (StringUtils.isBlank(contentType)) {
            contentType = DEFAULT_CONTENT_TYPE;
        }
        method.addHeader("Content-Type", contentType);
        method.addHeader("Accept", contentType);
        method.setConfig(requestConfig);
        return method;
    }

    /**
     * 执行GET 请求
     *
     * @param uri
     * @return
     */
    public static String execute(String uri) {
        long startTime = System.currentTimeMillis();
        HttpEntity httpEntity = null;
        HttpRequestBase method = null;
        String responseBody = "";
        try {
            if (httpClient == null) {
                initPools();
            }
            method = getRequest(uri, HttpGet.METHOD_NAME, DEFAULT_CONTENT_TYPE, 0);
            HttpContext context = HttpClientContext.create();
            CloseableHttpResponse httpResponse = httpClient.execute(method, context);
            httpEntity = httpResponse.getEntity();
            if (httpEntity != null) {
                responseBody = EntityUtils.toString(httpEntity, "UTF-8");
             //   logger.info("请求URL: "+uri+"+  返回状态码："+httpResponse.getStatusLine().getStatusCode());
            }
        } catch (Exception e) {
            if(method != null){
                method.abort();
            }
            e.printStackTrace();
          /*  logger.error("execute get request exception, url:" + uri + ", exception:" + e.toString() + ",cost time(ms):"
                    + (System.currentTimeMillis() - startTime));*/
        } finally {
            if (httpEntity != null) {
                try {
                    EntityUtils.consumeQuietly(httpEntity);
                } catch (Exception e) {
                    e.printStackTrace();
                //    logger.error("close response exception, url:" + uri + ", exception:" + e.toString() + ",cost time(ms):"
                  //          + (System.currentTimeMillis() - startTime));
                }
            }
        }
        return responseBody;
    }

    public static void main(String[] args) throws ParseException {
    //    String token = "https://sharpeyespc.zbj.com/access/account/110054801/secret/SHW2sf4WfrTxc";
        String s ="https://sharpeyespc.zbj.com/enterprise/information?companyName=磐安县海润塑胶厂&type=1&access_token=9d646c6f45e3e1f7724055461ade4ffc";
        String sd =  HttpClientPoolUtil.execute(s);
    }
}
