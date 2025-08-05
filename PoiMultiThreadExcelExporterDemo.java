package org.example;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class PoiMultiThreadExcelExporterDemo {

    // 示例：导出用户数据
    public static void main(String[] args) {
        // 1. 定义Excel表头
        String[] headers = {"ID", "姓名", "年龄", "邮箱", "部门", "入职日期"};

        // 2. 准备模拟数据（10万条记录，用于测试大数据量导出）
        List<Employee> dataList = generateTestData(100000);

        // 3. 创建导出器实例
        // 数据映射：将Employee对象转换为Excel行数据（与表头顺序对应）
        PoiMultiThreadExcelExporter<Employee> exporter = new PoiMultiThreadExcelExporter<>(
            headers,
            employee -> new Object[]{
                employee.getId(),
                employee.getName(),
                employee.getAge(),
                employee.getEmail(),
                employee.getDepartment(),
                employee.getHireDate()
            },
            "员工信息表",  // Sheet基础名称
            5,  // 线程池大小（可选，默认是CPU核心数+1）
            null,  // 表头样式（可选，使用默认）
            null   // 单元格样式（可选，使用默认）
        );

        try {
            // 4. 执行导出
            // 每个Sheet存放10000条数据，最终会生成10个Sheet（100000/10000）
            exporter.export(dataList, "D:/员工信息汇总.xlsx", 10000);
            System.out.println("Excel导出成功！文件路径：D:/员工信息汇总.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
            System.err.println("导出失败：" + e.getMessage());
        }
    }

    // 生成模拟数据
    private static List<Employee> generateTestData(int count) {
        List<Employee> list = new ArrayList<>(count);
        for (int i = 0; i < count; i++) {
            Employee emp = new Employee();
            emp.setId(i + 10000L);  // ID从10000开始
            emp.setName("员工" + (i + 1));
            emp.setAge(22 + (i % 30));  // 年龄22-51岁
            emp.setEmail("emp" + (i + 1) + "@company.com");
            emp.setDepartment(getDepartment(i % 5));  // 5个部门循环
            emp.setHireDate(new Date(System.currentTimeMillis() - (i * 86400000L * 30)));  // 模拟入职日期
            list.add(emp);
        }
        return list;
    }

    // 部门名称映射
    private static String getDepartment(int index) {
        switch (index) {
            case 0: return "技术部";
            case 1: return "市场部";
            case 2: return "财务部";
            case 3: return "人力资源部";
            case 4: return "运营部";
            default: return "其他";
        }
    }

    // 数据实体类：员工信息
    static class Employee {
        private Long id;
        private String name;
        private Integer age;
        private String email;
        private String department;
        private Date hireDate;

        // Getter方法（必须实现，用于数据映射）
        public Long getId() { return id; }
        public void setId(Long id) { this.id = id; }

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }

        public Integer getAge() { return age; }
        public void setAge(Integer age) { this.age = age; }

        public String getEmail() { return email; }
        public void setEmail(String email) { this.email = email; }

        public String getDepartment() { return department; }
        public void setDepartment(String department) { this.department = department; }

        public Date getHireDate() { return hireDate; }
        public void setHireDate(Date hireDate) { this.hireDate = hireDate; }
    }
}