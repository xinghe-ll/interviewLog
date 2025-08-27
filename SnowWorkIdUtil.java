/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2024-2024. All rights reserved.
 */

package com.huawei.fin.bfd.goams.domain.common.util;

import com.huawei.it.jalor5.core.util.StringUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.InetAddress;
import java.security.SecureRandom;
import java.util.UUID;

/**
 * 基于twitter的雪花算法
 *
 * @author z00678240
 * @since 2023-08-09
 */
public final class SnowWorkIdUtil {
    private static final Logger logger = LoggerFactory.getLogger(SnowWorkIdUtil.class);

    /**
     * 时间起始标记点，作为基准，一般取系统的最近时间
     */
    private static final long EPOCH_NUM = 1618541381557L;

    /**
     * 机器标识位数
     */
    private static final long WORKER_IDBITS = 10L;

    /**
     * 机器ID最大值: 1023
     */
    private static final long MAXWORKER_ID = -1L ^ -1L << WORKER_IDBITS;

    /**
     * 毫秒内自增位
     */
    private static final long SEQUENCE_BITS = 12L;

    /**
     * 12
     */
    private static final long WORKER_IDSHIFT = SEQUENCE_BITS;

    /**
     * 22
     */
    private static final long TIMESTAMP_LEFTSHIFT = WORKER_IDBITS + SEQUENCE_BITS;

    /**
     * 4095,111111111111,12位
     */
    private static final long SEQUENCE_MASK = -1L ^ -1L << SEQUENCE_BITS;

    /**
     * 每台机器分配不同的id
     */
    private static long iworker;

    /**
     * 0，并发控制
     */
    private static long sequenceSe = 0L;

    private static long lastTimestamp = -1L;

    static {
        String hostIp;
        try {
            InetAddress netAddress = InetAddress.getLocalHost();
            hostIp = netAddress.getHostAddress();
            if (!StringUtil.isNullOrEmpty(hostIp)) {
                StringBuilder hostNo = new StringBuilder();
                for (int i = 0; i < hostIp.length(); i++) {
                    if (hostIp.charAt(i) >= 48 && hostIp.charAt(i) <= 57) {
                        // 取最后一组数字
                        hostNo.append(hostIp.charAt(i));
                    }
                }
                if (!"".equals(hostNo.toString())) {
                    iworker = Long.parseLong(hostNo.toString()) % MAXWORKER_ID;
                } else {
                    iworker = new SecureRandom().nextInt(1000) + 1;
                }
            } else {
                iworker = new SecureRandom().nextInt(1000) + 1;
            }
        } catch (Exception exception) {
            logger.error("iWorker ID init not ok", exception);
            iworker = new SecureRandom().nextInt(1000) + 1;
        }
    }

    /**
     * Next id long
     *
     * @return the long
     */
    public static synchronized long getNextId() {
        long timestamp = timeGen();
        // 如果上一个timestamp与新产生的相等，则sequence加一(0-4095循环); 对新的timestamp，sequence从0开始
        if (lastTimestamp == timestamp) {
            sequenceSe = sequenceSe + 1 & SEQUENCE_MASK;
            if (sequenceSe == 0L) {
                // 重新生成timestamp
                timestamp = tilNextMillis(lastTimestamp);
            }
        } else {
            sequenceSe = 0L;
        }
        if (timestamp < lastTimestamp) {
            UUID uuid = UUID.randomUUID();
            return uuid.getMostSignificantBits();
        }
        lastTimestamp = timestamp;
        return timestamp - EPOCH_NUM << TIMESTAMP_LEFTSHIFT | iworker << WORKER_IDSHIFT | sequenceSe;
    }

    /**
     * 获取下一个ID
     *
     * @return long
     */
    public static synchronized long nextId() {
        return getNextId();
    }

    /**
     * 获取下一个ID
     *
     * @return string
     */
    public synchronized String nextIdString() {
        return String.valueOf(getNextId());
    }

    /**
     * 等待下一个毫秒的到来, 保证返回的毫秒数在参数lastTimestamp之后
     *
     * @param lastTimestamp lastTimestamp
     * @return long timestamp
     */
    private static long tilNextMillis(long lastTimestamp) {
        long timestamp = timeGen();
        while (timestamp <= lastTimestamp) {
            timestamp = timeGen();
        }
        return timestamp;
    }

    /**
     * 获得系统当前毫秒数
     *
     * @return the long
     */
    public static long timeGen() {
        return System.currentTimeMillis();
    }
}
