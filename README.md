# interviewLog
面试总结 2025-06-14
___
# spring
1. spring相比非spring项目有优点
2. spring的IOC, AOP
3. spring事务管理
4. springboot关键注解
5. springboot的自动配置原理
6. springboot使用了哪些设计模式
7. springcloud核心组件(nacos,ribbon,openfeign,hystrix/sentinel,zuul/gateway)
8. RPC原理?使用哪个组件?
9. ai工具使用?
10. K8S, docker
11. springMVC流程
12. 过滤器, 拦截器, 全局统一异常处理
13. 浏览器输入url后发生了啥?


# 数据库
1. 多jvm同时操作同一张表同一行数据,怎么保证数据安全?
->行锁(例:where id=? for udpate)
->乐观锁: 数据行加版本号控制(例: update .. where version=?)
2. sql性能分析,性能优化,索引失效场景
3. 乐观锁,悲观锁的理解
4. CAS思想
5. mysql的索引结构(B+树),走索引为什么快
6. mysql的mvcc
7. 事务ACID特性: 原子性,一致性,隔离性,持久性
8. 事务隔离级别(读未提交,读已提交,可重复读,串行化)
9. 纯查询操作是否需要事务?
->为保证多表时间切片一致性,需要事务
->先查后改的场景,需要事务
->纯粹简单查询,不需要事务
10. 分库分表?
11. 分布式事务?
12. 多数据源配置?

# 缓存
1. 本地缓存方案?
2. redis缓存
3. redis(集群哨兵模式, cluster模式)
4. redis持久化
5. redis分布式锁
6. 分布式锁实现方式(redisson, 数据库唯一主键索引(不常用),zookepper)

# mq
1. 实际项目mq用在哪些地方?
2. 消息幂等性保证
3. 消息重复消费
4. 消息顺序消费
5. 怎么保证消息不丢失
6. mq消息消费均匀分布?
7. p2p,广播模式
8. kafka?


# java基础
1. java.util.concurrent包下的常用类
2. 线程池配置实现动态调整参数
3. OOM分析定位
4. CPU100%分析定位
5. jvm内存结构
7. jvm垃圾回收
8. java类加载器,双亲委派机制
9. java泛型
10. java反射
11. APM监控工具
12. elk使用
13. 权限模型？？(RBAC  ABAC等)
____

// update nothing.



