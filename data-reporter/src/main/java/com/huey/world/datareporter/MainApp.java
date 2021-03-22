package com.huey.world.datareporter;

import com.huey.world.datareporter.datacp.WorkbookContext;
import com.huey.world.datareporter.datacp.worker.CpWorker;
import com.huey.world.datareporter.datacp.worker.DateWorker;
import com.huey.world.datareporter.datacp.worker.SortWorker;
import com.huey.world.datareporter.datacp.worker.Workable;
import lombok.extern.slf4j.Slf4j;

import java.util.stream.Stream;

/**
 * @author huey
 */
@Slf4j
public class MainApp {

    public static void main(String[] args) {

        String cfgPath = System.getProperty("cfg.file");
        WorkbookContext context = new WorkbookContext(cfgPath);

        Workable[] workers = new Workable[] {
                new CpWorker(context),
                new DateWorker(context),
                new SortWorker(context)
        };

        Stream.of(workers).forEach(Workable::execute);

        context.exit();

    }

}
