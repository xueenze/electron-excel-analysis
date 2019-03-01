var xlsx = require('xlsx');
var fs = require('fs');
var shell = require('electron').shell;

(function() {
    /**
     * excel处理对象
     * @param {*} inputSource 
     * @param {*} outputSource 
     */
    function operateExcel(inputSource, outputSource) {
        this.inputSource = inputSource;
        this.outputSource = outputSource;
        this.excelJsonData = {};
        this.succeedOutput = 0;
    }

    /**
     * 第一步
     * 判断文件路径和输出目录的有效性
     */
    operateExcel.prototype.firstStep = function() {
        try {
            fs.accessSync(this.inputSource, fs.F_OK);
            fs.accessSync(this.outputSource, fs.F_OK);

            console.log(this.inputSource);
            console.log(this.outputSource);

            return true;
        } catch(e) {
            return false;
        }
    }

    /**
     * 获取所有的列名
     */
    operateExcel.prototype.previewExcel = function() {
        let workbook = xlsx.readFile(this.inputSource);
        let worksheet = workbook.Sheets[workbook.SheetNames[0]];

        this.excelJsonData = xlsx.utils.sheet_to_json(worksheet);

        if (this.excelJsonData.length > 0) {
            // 如果有数据的话，取出第一条数据即可
            let columns = Object.getOwnPropertyNames(this.excelJsonData[0]);
            columns.shift();
            
            const columnNames = $('#columnsNames');

            columns.forEach(name => {
                columnNames.append(`<span class="label label-info">${name}</span>`);
            });
        }
    }

    /**
     * 处理Excel
     */
    operateExcel.prototype.operateExcel = function() {
        if (this.excelJsonData.length > 0) {
            let sum = this.excelJsonData.length;

            this.excelJsonData.forEach(item => {
                let fileName = `${item['序号']}-${item['姓名']}`;
                let password = item['密码'];
                delete item['密码'];
    
                let wb = { SheetNames: ['Sheet1'], Sheets: {}, Props: {} };
                wb.Sheets['Sheet1'] = xlsx.utils.json_to_sheet([item]);

                try {
                    xlsx.writeFile(wb, `${this.outputSource}/${fileName}.xlsx`);
                    this.succeedOutput++;

                    $('#progressBar').text(`${(this.succeedOutput * 100 / sum)}%`);

                    $('#progressBar').css(
                        'width', 
                        `${(this.succeedOutput * 100 / sum)}%`
                    );
                } catch(e) {
                    console.log(e);
                }
            });
        }
    }

    /**
     * 初始化所有事件
     */
    function init() {
        var operateItem = {};

        const step1 = $('#step1');
        step1.bind('click', function() {
            let inputSource = document.getElementById('inputSource').value;
            let outputSource = document.getElementById('outputSource').value;
            
            // 将第一步的路径信息传入第二步中
            let step3OutputSource = document.getElementById('step3-output-source');

            step3OutputSource.innerHTML = outputSource + '/';

            operateItem = new operateExcel(inputSource, outputSource);

            if (operateItem.firstStep()) {
                operateItem.previewExcel();
                step1.removeClass('active');
                step2.addClass('active');
            }
        });

        const step2 = $('#step2');
        const step3 = $('#step3');;
        step2.bind('click', function() {
            operateItem.operateExcel();
            step2.removeClass('active');
            step3.addClass('active');
        });

        const step3OutputSource = document.getElementById('step3-output-source');
        step3OutputSource.addEventListener('click', function(){
            shell.showItemInFolder(step3OutputSource.innerText);
        });
    }

    init();
})();