# 成绩单工程
Python 学习用
- 该工程总文件夹名为all_docs
  - 子文件夹名以A2204开头，为虚构的学号
    - 每个文件夹内包含3个.docx文档，分别为：
      - 指导教师评价表
        - 包含开题、外文、设计、创新、撰写、态度、综合评分
      - 评阅评语表
        - 包含开题、外文、设计、创新、撰写、综合评分
      - 成绩表
        - 包含设计、创新、答辩、综合评分
    - 每个.docx文件名都以“学号_学生姓名_”开头

要求：
1. 自行查阅能够处理office文档python资源，选取合适的库，编写python代码，使其能一键自动生成一份excel表格（样式格式可参考ref.xlsx文件），该表含每位同学的基本信息（姓名、学号、专业、指导老师、毕设题目）以及7+6+4共17项分数
2. 表格包含清晰的“表头”以示每项数据的含义，视情况合并部分单元格
3. 除所有学生的毕业设计题目文字为居左，其他所有单元格文字皆居中
4. 为表格内容加边框
