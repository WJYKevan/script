const puppeteer = require('puppeteer');
const readline = require('readline');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const DOWNLOAD_DIR = 'D:\\Users\\dh\\Desktop\\测试数据\\';
const RESET = '\x1b[0m'; // 重置所有属性
const GREEN = '\x1b[32m'; // 绿色
const RED = '\x1b[31m'; // 红色


(async () => {
    try {
        const browser = await puppeteer.launch({ headless: false });
        const page = await browser.newPage();
        console.log(`\n-------------------------------------------------------------------------------------------------------`);
        console.log(`                       1. 浏览器启动成功，开始执行脚本...                                        `);
        console.log(`-------------------------------------------------------------------------------------------------------`);
        // 登录过程
        await page.goto('https://supply.zte.com.cn/UI/Web/Application/kxscm/kxsup_manager/Portal/index.aspx');
        await page.waitForSelector('#txt_userno');
        await page.type('#txt_userno', 'duanhua');
        await page.type('#txt_pwd', 'Fh2865358');
        await page.click('#chb_privacy_policy');
        console.log(`请输入验证码后按回车继续...\n`);
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });
        const captchaPromise = new Promise(resolve => rl.question(` 验证码: `, resolve));
        const captchaInput = await captchaPromise;
        rl.close();
        await page.type('#txt_veriCode', captchaInput);
        await page.click('#btn_login');
        await page.waitForNavigation({ waitUntil: 'networkidle2' });
        console.log(`\n-------------------------------------------------------------------------------------------------------`);
        console.log(`                       2. 登录成功，跳转至目标页面...                                         `);
        console.log(`-------------------------------------------------------------------------------------------------------`);
        await page.goto('https://isrm.zte.com.cn/zte-iss-bobase-portalui/#/app/ztea_iSRMW_external/page/ztef_iSRMG_chzlsjwh');
        // 确保iframe加载完成
        await page.waitForSelector('iframe[src*="zte-isrm-itemquality-fe"]', { visible: true });
        // 获取iframe
        const iframes = await page.$$('iframe');
        let targetFrame;
        for (const iframe of iframes) {
            const src = await page.evaluate(element => element.src, iframe);
            if (src.includes('zte-isrm-itemquality-fe')) {
                targetFrame = iframe;
                break;
            }
        }
        if (targetFrame) {
            const frame = await targetFrame.contentFrame();
            // 在iframe中查找并点击按钮
            const buttonText = '新增';
            const button = await frame.waitForFunction(buttonText => {
                const buttons = document.querySelectorAll('button');
                return Array.from(buttons).find(btn => btn.innerText.trim() === buttonText);
            }, {}, buttonText);

            if (button) {
                await button.click();
            } else {
                console.error('在iframe中未找到匹配的按钮');
            }
        } else {
            iframes.forEach(async (iframe, index) => {
                const src = await page.evaluate(element => element.src, iframe);
                console.log(`iframe #${index + 1}: ${src}`);
            });
        }
        // console.log(`\n`);
        console.log(`                       3. 正在预设路径${RED}D:/Users/dh/Desktop/测试数据/${RESET}下寻找未处理文件...   `);
        console.log(`-------------------------------------------------------------------------------------------------------`);
        // 查找未被处理的Excel文件
        let unprocessedFile, boxNumber;
        fs.readdirSync(DOWNLOAD_DIR).forEach(file => {
            const fileNameWithoutExt = path.parse(file).name;
            const ext = path.extname(file);
            if (ext === '.xlsx' && !fileNameWithoutExt.endsWith('(已维护)')) {
                unprocessedFile = file;
                boxNumber = fileNameWithoutExt; // 保存原始文件名作为箱号
                return; // 找到第一个符合条件的文件后即停止循环
            }
        });
        if (unprocessedFile) {
            const originalFilePath = path.join(DOWNLOAD_DIR, unprocessedFile);
            const processedFileName = `${boxNumber}(已维护).xlsx`; // 在文件名末尾加"(已维护)"
            const processedFilePath = path.join(DOWNLOAD_DIR, processedFileName);
            // 读取Excel文件
            const workbook = xlsx.readFile(originalFilePath);
            const sheetName = workbook.SheetNames[0];
            const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
            const firstRow = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName])[0]; // 取第一行数据
            orderNumber = firstRow['订单号'].toString();
            // 提取订单号最后五位数字
            const lastFiveDigits = orderNumber.slice(-5);
            // 根据7,8位数字确定字母（映射A-Z）
            const combinedDigits = parseInt(`${orderNumber[6]}${orderNumber[7]}`, 10);
            const letterFrom7and8 = String.fromCharCode('A'.charCodeAt(0) + (combinedDigits - 1)); // 计算对应的ASCII码并转换为字符
            // 根据3到6位数字映射字母（这里假设是年份，简化处理为固定映射）
            const yearDigits = orderNumber.substring(2, 6);
            let letterFromYear;
            switch (yearDigits) {
                case '2024':
                    letterFromYear = 'U'; // 示例映射，您可以根据实际情况调整
                    break;
                // 可以添加更多case来处理其他年份
                default:
                    letterFromYear = ''; // 或者设定一个默认值
            }
            // 组合成批次号
            batchNumber = `${letterFromYear}${letterFrom7and8}${lastFiveDigits}`;
            console.log(` 已经找到箱号: ${GREEN}${boxNumber}${RESET}, 批次号: ${GREEN}${batchNumber}${RESET} 的文件，正在处理中...\n`);
            // 从Excel读取并转换后的数据
            const sequenceAndMaterialData = data.map(row => {
                // 提取序号（第一列）
                const sequence = row['序号'];
                // 构造生产批次，批次号后加上序号，序号不足三位前面补0
                const productionBatch = `${batchNumber}${sequence.toString().padStart(3, '0')}`;
                // 提取客户料号（第五列），并确保转换为12位数字，不足部分在前面补零
                const materialNumberRaw = row['客户料号'];
                const materialNumber = String(materialNumberRaw).padStart(12, '0');
                // 返回一个包含序号和格式化后的客户物料号的对象
                return { productionBatch, materialNumber };
            });
            // // 示例输出生产批次和客户料号信息
            // sequenceAndMaterialData.forEach(({ productionBatch, materialNumber }) => {
            //     console.log(`生产批次: ${GREEN}${productionBatch}${RESET}, 客户物料号: ${GREEN}${materialNumber}${RESET}`);
            // });

            const { materialNumber } = sequenceAndMaterialData[0];


            console.log(`物料号: ${materialNumber} 已自动填入输入框。`);

            // 处理完文件后，更改文件名
            // fs.renameSync(originalFilePath, processedFilePath);
            // console.log(`文件处理完成，已重命名为: ${processedFilePath}`);

            // 遍历sequenceAndMaterialData中的每条记录
            for (const { materialNumber } of sequenceAndMaterialData) {
                // 确保回到了主frame
                await page.bringToFront();

                // 再次获取目标iframe
                const iframes = await page.$$('iframe');
                let currentTargetFrame = iframes.find(async iframe =>
                    (await page.evaluate(element => element.src, iframe)).includes('zte-isrm-itemquality-fe')
                );

                if (currentTargetFrame) {
                    const frame = await currentTargetFrame.contentFrame();

                    // 在iframe中查找物料代码输入框并填入物料号
                    await frame.evaluate((materialNumber) => {
                        const materialCodeInput = document.querySelector('input[type="text"][autocomplete="off"][placeholder="请选择物料代码"]');
                        if (materialCodeInput) {
                            materialCodeInput.value = materialNumber;
                            // 触发input事件
                            const event = new Event('input', { bubbles: true });
                            materialCodeInput.dispatchEvent(event);
                        } else {
                            console.error('未能找到物料代码输入框');
                        }
                    }, materialNumber);

                    console.log(`物料号: ${materialNumber} 已自动填入输入框。`);
                } else {
                    console.error('在重新查找时未找到目标iframe');
                }
            }


        } else {
            console.log('没有找到未处理的Excel文件。');
        }






    } catch (error) {
        console.error('执行过程中发生错误:', error);
    }
})();