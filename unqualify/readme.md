# 不合格品管理系统

## 项目说明
这是一个处理不合格品管理的系统，主要功能包括：
1. 项目编号关联管理
2. 工序编号关联管理 
3. 现象、模式和原因的级联选择

## 功能说明
### 1. 项目编号关联
- 当选择项目编号时，自动加载对应的车号选项
- 车号支持多选，默认包含"All"选项

### 2. 工序编号关联
- 当选择工序编号时，自动加载对应的现象选项

### 3. 级联选择
- 选择现象后，自动加载对应的模式选项
- 选择模式后，自动加载对应的原因选项

## 技术实现
- 使用JavaScript ES6+
- 采用面向对象的方式组织代码
- 实现数据与UI的分离 