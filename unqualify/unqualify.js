/**
 * 不合格品管理插件
 */
define(['bili/AbstractBillPlugin'], function(AbstractBillPlugin) {
    return {
        extend: AbstractBillPlugin,
        
        /**
         * 属性改变事件处理
         * @param {Object} e - 事件参数
         */
        propertyChanged: function(e) {
            this._super();
            
            // 处理项目编号关联
            this.handleProjectNumberChange();
            
            // 处理工序编号关联
            this.handleProcessNumberChange();
            
            // 处理级联选择
            this.handleCascadeSelection();
        },
        
        /**
         * 数据加载后事件处理
         * @param {Object} e - 事件参数
         */
        afterBindData: function(e) {
            this._super();
            
            // 处理项目编号关联
            this.handleProjectNumberChange();
            
            // 处理工序编号关联
            this.handleProcessNumberChange();
            
            // 处理级联选择
            this.handleCascadeSelection();
        },
        
        /**
         * 处理项目编号变化
         */
        handleProjectNumberChange: function() {
            var projectObj = this.getModel().getValue("crrc_projnum");
            if (!projectObj) return;
            
            var project = this.loadSingle(projectObj.getPkValue(), "crrc_project1");
            var jchs = project.getString("crrc_textfield");
            
            if (!this.isEmpty(jchs)) {
                var items = jchs.split(",");
                var jchItems = this.getControl("crrc_jcnum");
                var datas = [];
                
                // 添加All选项
                datas.push({
                    text: "All",
                    value: "All"
                });
                
                // 添加其他选项
                items.forEach(function(item) {
                    datas.push({
                        text: item,
                        value: item
                    });
                });
                
                jchItems.setComboItems(datas);
            }
        },
        
        /**
         * 处理工序编号变化
         */
        handleProcessNumberChange: function() {
            var gxnumber = this.getModel().getValue("crrc_gxnumber");
            if (!gxnumber) return;
            
            var gxx = this.loadSingle(gxnumber.getPkValue(), "crrc_gxx");
            var phenom = gxx.getString("crrc_phenom1");
            
            if (!this.isEmpty(phenom)) {
                var phenoms = phenom.split(",");
                var phenomItems = this.getControl("crrc_phenom1");
                var datas = [];
                
                phenoms.forEach(function(item) {
                    datas.push({
                        text: item,
                        value: item
                    });
                });
                
                phenomItems.setComboItems(datas);
            }
        },
        
        /**
         * 处理级联选择
         */
        handleCascadeSelection: function() {
            var phenom = this.getModel().getValue("crrc_phenom1");
            if (!phenom) return;
            
            var filter = [
                ["crrc_phenom1", "=", phenom]
            ];
            
            var phenomRecords = this.load("crrc_xxkkx", [
                "name",
                "crrc_xxkkx_entryentity.crrc_mode",
                "crrc_xxkkx_entryentity.crrc_xxkkx_subentryentity",
                "crrc_xxkkx_subentryentity.crrc_cause"
            ], filter);
            
            // 更新模式选项
            this.updateModeItems(phenomRecords);
            
            // 如果已选择模式，更新原因选项
            var mode = this.getModel().getValue("crrc_mode");
            if (mode) {
                this.updateCauseItems(phenomRecords);
            }
        },
        
        /**
         * 更新模式选项
         * @param {Array} phenomRecords - 现象记录
         */
        updateModeItems: function(phenomRecords) {
            var modeItems = this.getControl("crrc_mode");
            var modeDatas = [];
            
            phenomRecords.forEach(function(record) {
                var modeEntries = record.get("crrc_xxkkx_entryentity");
                modeEntries.forEach(function(entry) {
                    var mode = entry.getString("crrc_mode");
                    modeDatas.push({
                        text: mode,
                        value: mode
                    });
                });
            });
            
            modeItems.setComboItems(modeDatas);
        },
        
        /**
         * 更新原因选项
         * @param {Array} phenomRecords - 现象记录
         */
        updateCauseItems: function(phenomRecords) {
            var causeItems = this.getControl("crrc_cause");
            var causeDatas = [];
            
            phenomRecords.forEach(function(record) {
                var modeEntries = record.get("crrc_xxkkx_entryentity");
                modeEntries.forEach(function(modeEntry) {
                    var causeEntries = modeEntry.get("crrc_xxkkx_subentryentity");
                    causeEntries.forEach(function(causeEntry) {
                        var cause = causeEntry.getString("crrc_cause");
                        causeDatas.push({
                            text: cause,
                            value: cause
                        });
                    });
                });
            });
            
            causeItems.setComboItems(causeDatas);
        },
        
        /**
         * 判断字符串是否为空
         * @param {string} str - 待检查的字符串
         * @returns {boolean} 是否为空
         */
        isEmpty: function(str) {
            return !str || str.trim().length === 0;
        },
        
        /**
         * 加载单条记录
         * @param {string} pkValue - 主键值
         * @param {string} formId - 表单ID
         * @returns {Object} 记录对象
         */
        loadSingle: function(pkValue, formId) {
            return $data.load({
                billFormId: formId,
                pkValue: pkValue
            })[0];
        },
        
        /**
         * 加载记录集合
         * @param {string} formId - 表单ID
         * @param {Array} fields - 字段列表
         * @param {Array} filter - 过滤条件
         * @returns {Array} 记录集合
         */
        load: function(formId, fields, filter) {
            return $data.load({
                billFormId: formId,
                fields: fields,
                filter: filter
            });
        }
    };
}); 