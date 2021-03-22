package com.huey.world.datareporter.datacp.cfg;

import com.alibaba.fastjson.JSON;
import com.huey.world.datareporter.common.exception.DataReporterException;
import com.huey.world.datareporter.datacp.cfg.spec.Cfg;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;

/**
 * @author huey
 */
public class CfgReader {

    /**
     * 读取配置文件
     *
     * @param cfgPath
     * @return
     * @throws IOException
     */
    public Cfg readCfg(String cfgPath) {

        // 读取配置文件内容
        File cfgFile = new File(cfgPath);
        String cfgText = null;
        try {
            cfgText = FileUtils.readFileToString(cfgFile, Charset.defaultCharset());
        }
        catch (IOException e) {
            throw new DataReporterException(e);
        }

        return JSON.parseObject(cfgText, Cfg.class);

    }

}
