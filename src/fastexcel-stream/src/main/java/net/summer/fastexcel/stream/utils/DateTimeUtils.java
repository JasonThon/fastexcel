package net.summer.fastexcel.stream.utils;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;

/**
 * @author songyi
 * @Date 2021/12/4
 **/
public class DateTimeUtils {
    public static LocalDateTime toLocal(final Instant instant) {
        return LocalDateTime.ofInstant(instant, ZoneId.systemDefault());
    }
}
