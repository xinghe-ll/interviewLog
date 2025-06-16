# interviewLog
面试总结 2025-06-14
___
# spring
spring相比非spring项目有优点
spring的IOC, AOP
spring事务管理
springboot关键注解
springboot的自动配置原理
springboot使用了哪些设计模式
springcloud核心组件(nacos,ribbon,openfeign,hystrix/sentinel,zuul/gateway)
RPC原理?使用哪个组件?
ai工具使用?
K8S, docker
springMVC流程
过滤器, 拦截器, 全局统一异常处理
浏览器输入url后发生了啥?


# 数据库
sql性能分析,性能优化,索引失效场景
多jvm同时操作同一张表同一行数据,怎么保证数据安全?
->行锁(例:where id=? for udpate)
->乐观锁: 数据行加版本号控制(例: update .. where version=?)
乐观锁,悲观锁的理解
CAS思想
mysql的索引结构(B+树),走索引为什么快
mysql的mvcc
事务ACID特性: 原子性,一致性,隔离性,持久性
事务隔离级别(读未提交,读已提交,可重复读,串行化)
纯查询操作是否需要事务?
->为保证多表时间切片一致性,需要事务
->先查后改的场景,需要事务
->纯粹简单查询,不需要事务
分库分表?
分布式事务?
多数据源配置?

# 缓存
本地缓存方案?
redis缓存
redis(集群哨兵模式, cluster模式)
redis持久化
redis分布式锁
分布式锁实现方式(redisson, 数据库唯一主键索引(不常用),zookepper)

# mq
实际项目mq用在哪些地方?
消息幂等性保证
消息重复消费
消息顺序消费
怎么保证消息不丢失
mq消息消费均匀分布?
p2p,广播模式
kafka?


# java基础
java.util.concurrent包下的常用类
线程池配置实现动态调整参数
OOM分析定位
CPU100%分析定位
jvm内存结构
jvm垃圾回收
java类加载器,双亲委派机制
java泛型
java反射
APM监控工具
elk使用
权限模型？？(RBAC  ABAC等)
____

// update nothing.



