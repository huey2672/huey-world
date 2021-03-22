package com.huey.world.datareporter.common.exception;

/**
 * @author huey
 */
public class DataReporterException extends RuntimeException {

    public DataReporterException() {
        super();
    }

    public DataReporterException(String message) {
        super(message);
    }

    public DataReporterException(String message, Throwable cause) {
        super(message, cause);
    }

    public DataReporterException(Throwable cause) {
        super(cause);
    }

}
