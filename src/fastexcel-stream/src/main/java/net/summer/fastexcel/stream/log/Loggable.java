package net.summer.fastexcel.stream.log;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author songyi
 * @Date 2021/12/4
 **/

public interface Loggable {

    /**
     * @param message
     * @param arguments
     * @see Logger#debug(String, Object...)
     */
    default void debug(final String message, Object... arguments) {
        if (getLogger().isDebugEnabled()) {
            getLogger().debug(message, arguments);
        }
    }

    /**
     * @param message
     * @param throwable
     * @see Logger#debug(String, Throwable)
     */
    default void debug(final String message, Throwable throwable) {
        if (getLogger().isDebugEnabled()) {
            getLogger().debug(message, throwable);
        }
    }

    /**
     * @param message
     * @param arguments
     * @see Logger#error(String, Object...)
     */
    default void error(final String message, Object... arguments) {
        if (getLogger().isErrorEnabled()) {
            getLogger().error(message, arguments);
        }
    }

    /**
     * @param message
     * @param throwable
     * @see Logger#error(String, Throwable)
     */
    default void error(final String message, Throwable throwable) {
        if (getLogger().isErrorEnabled()) {
            getLogger().error(message, throwable);
        }
    }

    default Logger getLogger() {
        return LoggerFactory.getLogger(this.getClass());
    }

    /**
     * @param message
     * @param arguments
     * @see Logger#info(String, Object...)
     */
    default void info(final String message, Object... arguments) {
        if (getLogger().isInfoEnabled()) {
            getLogger().info(message, arguments);
        }
    }

    /**
     * @param message
     * @param arguments
     * @see Logger#warn(String, Object...)
     */
    default void warn(final String message, Object... arguments) {
        if (getLogger().isWarnEnabled()) {
            getLogger().warn(message, arguments);
        }
    }

    /**
     * @param message
     * @param throwable
     * @see Logger#warn(String, Throwable)
     */
    default void warn(final String message, Throwable throwable) {
        if (getLogger().isWarnEnabled()) {
            getLogger().warn(message, throwable);
        }
    }
}

