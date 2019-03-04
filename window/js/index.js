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
            columnNames.empty();

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
                let fileName = `${item['身份证']}-${item['姓名']}`;
                let password = item['密码'];
                delete item['密码'];
    
                let wb = { SheetNames: ['Sheet1'], Sheets: {}, Props: {} };
                wb.Sheets['Sheet1'] = xlsx.utils.json_to_sheet([item]);

                try {
                    xlsx.writeFile(wb, `${this.outputSource}/${fileName}.xlsx`);
                    this.succeedOutput++;
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
        const step2Prev = $('#step2-prev');
        const step2Next = $('#step2-next');
        const step3 = $('#step3');

        step2Prev.bind('click', function() {
            step1.addClass('active');
            step2.removeClass('active');

            $('#progressBar').css(
                'width', 
                '0%'
            );
        });

        step2Next.bind('click', () => {
            operateItem.operateExcel();

            // 这里加一个进度条的动画
            let count = 1;

            var handler = setInterval(function() {
                if (count > 10) {
                    clearInterval(handler);

                    step2.removeClass('active');
                    step3.addClass('active');
                } else {
                    $('#progressBar').css(
                        'width', 
                        `${count * 10}%`
                    );

                    $('#progressBar').text(`${count * 10}%`);

                    count++;
                }
            }, 200);

            $("#succeed-file-count").text(operateItem.succeedOutput);
        });

        const step3OutputSource = $('#step3-output-source');
        const step3Back = $('#step3-back');
        step3Back.bind('click', function() {
            step1.addClass('active');
            step3.removeClass('active');
        });
        step3OutputSource.bind('click', function(){
            shell.showItemInFolder(step3OutputSource.text());
        });
    }

    init();
})();