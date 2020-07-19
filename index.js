
const xlsx = require('node-xlsx')//xlsx 库
const fs = require('fs') //文件读写库
const request = require("request-promise-native");//request请求库
const cheerio = require("cheerio"); // 操作/解析Html  https://cheerio.js.org/   官方文档在这里，看不懂的自己嘤嘤嘤吧
let data = [] // 把这个数组写入excel   
request({
    url: "http://cep.ciosh.com/search/Search.aspx?k=&t=1&p=1&s=1050",//你要请求的地址
    method: "get",//请求方法 post get
    headers: {
        "content-type": "text/html",
        "Cookie": ""//如果携带了cookie
    },
}, async function (error, response, body) {
    if (!error && response.statusCode == 200) {
        $ = cheerio.load(body, {
            withDomLvl1: true,
            normalizeWhitespace: false,
            xmlMode: false,
            decodeEntities: true
        });
    
        let title = ['展商名称', '国家','省份', '产品品牌','展馆','展台','电话', '传真', '地址', '邮箱', '联系人']//设置表头
        data.push(title) // 添加完表头 下面就是添加真正的内容了
        let items = $(".c_result").children(".item");
     
        for (let i = 0; i < items.length; i++) {
            let arrInner = []
            let list=$(items[i]).children(".t").text();
            console.log(list);
            arrInner.push(list)
            arrInner.push($(items[i]).children(".c").text()?$(items[i]).children(".c").text().trim().split("省份")[0].trim().replace("国家：",""):"")
            arrInner.push($(items[i]).children(".c").text().trim().split("省份")[1]?$(items[i]).children(".c").text().trim().split("省份")[1].slice(1):"")
            arrInner.push($(items[i]).children(".d").text()?$(items[i]).children(".d").text().trim().slice(5):"")

            let id = $(items[i]).children(".t").children("a").attr("href").split("?")[1].split("=")[1]

            function  test(ms){
                let _request=request.get(`http://cep.ciosh.com/search/Exhibitor.aspx?eid=${id}`).then(response=>response).catch(err=>err)
                
                return new Promise((resolve, reject) => {
                    setTimeout(() => {
                        try{
                            resolve(_request)
                        }catch{
                            reject([])
                        } 
                    }, ms);
                })
            }

            let info = await test(0.1).catch(err=>{
                console.log('请求失败');
            }); 
            
            $info = cheerio.load(info, {
                withDomLvl1: true,
                normalizeWhitespace: false,
                xmlMode: false,
                decodeEntities: true
            });

            let table_info=$info(".OL_Booth table:nth-child(1) table:nth-child(1) tr:nth-child(1) td:nth-child(2)").text().trim();
         
            arrInner.push(table_info?table_info.split("展台号")[0].slice(3):"")
            arrInner.push(table_info?table_info.split("展台号:")[1]:"");
            arrInner.push($info(".OL_Contactus p:nth-child(1)").text().trim().slice(3))
            arrInner.push($info(".OL_Contactus p:nth-child(2)").text().trim().slice(3))
            arrInner.push($info(".OL_Contactus p:nth-child(3)").text().trim().slice(3))
            arrInner.push($info(".OL_Contactus p:nth-child(4)").text().trim().slice(3))
            arrInner.push($info(".OL_Contactus p:nth-child(5)").text().trim().slice(4))

            data.push(arrInner)
        }

        writeXls(data)
    }
});

process.on('unhandledRejection', error => {

    console.error('unhandledRejection', error);
    
    process.exit(1) // To exit with a 'failure' code
    
});


// 写xlsx文件
function writeXls(datas) {
    let buffer = xlsx.build([
        {
            name: 'sheet1',
            data: datas
        }
    ]);
    fs.writeFileSync('./xiaoreya.xlsx', buffer, { 'flag': 'w' });//生成excel data是excel的名字
}


