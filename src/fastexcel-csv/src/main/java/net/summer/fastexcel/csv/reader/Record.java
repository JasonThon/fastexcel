package net.summer.fastexcel.csv.reader;

import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.google.common.collect.Lists;

import net.summer.fastexcel.stream.utils.CollectionUtils;

import lombok.Getter;

/**
 * @author songyi
 * @Date 2021/12/1
 **/
@Getter
public class Record {
    private List<String> values;

    public String getItemAtIdx(final int currentIndex) {
        return CollectionUtils.safeGet(values, currentIndex);
    }

    public boolean isIdxOutOfBound(final int currentIndex) {
        return CollectionUtils.size(values) > currentIndex;
    }

    public Record newValues(final String line) {
        String[] strings = StringUtils.split(line, ",");
        if (this.values == null) {
            this.values = Lists.newArrayList();
        }

        this.values.clear();
        this.values.addAll(Lists.newArrayList(strings));

        return this;
    }
}
