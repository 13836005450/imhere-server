const Koa = require('koa')
const app = new Koa()
const url = '127.0.0.1/talent2'
const db = require('monk')(url)
const router = require('koa-router')()
const bodyParser = require('koa-bodyparser')
const http = require("http")
const https = require("https")
const fs = require("fs")
const static = require('koa-static')
const enforceHttps = require("koa-sslify").default
const path = require('path')
const cors = require('koa-cors')

const ejsexcel = require("ejsexcel");
const util = require("util");
const readFileAsync = util.promisify(fs.readFile);
const writeFileAsync = util.promisify(fs.writeFile);

const send = require('koa-send');

app.use(bodyParser())

async function generateGangWeiLeiBie2(arr){
    let newArr =arr;
    newArr.forEach(item => {
        let gangweileibie = item['info']["岗位类别"]
        let obj = {};
        gangweileibie.forEach(value => {
            let key = value['cas'].join('')
            let val = value['num'];
            obj[key] = val
        });
        item['info']['岗位类别2'] = obj;
    })
    return newArr
}
//深复制对象方法    
async function cloneObj(obj) {  
    let newObj = {};  
    if (obj instanceof Array) {  
        newObj = [];  
    }  
    for (let key in obj) {  
        let val = obj[key];  
        //newObj[key] = typeof val === 'object' ? arguments.callee(val) : val; //arguments.callee 在哪一个函数中运行，它就代表哪个函数, 一般用在匿名函数中。  
        newObj[key] = typeof val === 'object' ? cloneObj(val): val;  
    }  
    return newObj;  
}

async function generateExcel(data,excelName) {
    //获得Excel模板的buffer对象
    const exlBuf = await readFileAsync("excel/企业问卷20190818—横板—与电子问卷同.xlsx");
    //数据源
    const sourceData = data;
    //用数据源(对象)data渲染Excel模板
    const exlBuf2 = await ejsexcel.renderExcel(exlBuf, sourceData);
    await writeFileAsync("excel/"+excelName+".xlsx", exlBuf2);
    console.log("excel写入成功。。");
}

async function getUserId(user_name, user_password) {
    var collection = db.get('company_users')
    var user = await collection.findOne({ _user_name: user_name })
    if (!user) {
        return -1
    }
    if (user._user_password !== user_password) {
        return -2
    }
    return user._id
}

async function changeUserPass(user_name, user_id_card, user_new_pass) {
    var collection = db.get('company_users')
    var user = await collection.findOne({ _user_name: user_name})
    if (!user) {
        return -1
    }
    if (user._user_id_card !== user_id_card) {
        return -2
    }
    const update_result_user = await collection.update({
        _user_name: user_name
    }, {
            $set: {
                _user_password: user_new_pass
            }
        })
    return 1
}

async function getCompanyForm(user_id) {
    if (user_id === -1 || user_id === -2) {
        return -1
    } else {
        var collection = db.get('company_forms')
        var form = await collection.findOne({ _from_user: user_id.toString() })
        return form
    }
    //5d5564fc46bbe057fcaf56a0
    //5d5564fc46bbe057fcaf56a0
}

//ping!!!
router.get('/ping', async function (ctx, next) {
    ctx.body = 'pong!'
})

router.get('/new_user', async (ctx, next) => {
    var collections = db.get('forms')
    console.log('new user')
    var user_inserting = {
        _basic_status: false,
        _basic: {},
        _job_status: false,
        _job: {},
        _honor_status: false,
        _honor: {},
        _satisfaction_status: false,
        _satisfaction: {}
    }
    const result = await collections.insert(user_inserting)
    console.log(result)
    ctx.body = result._id
})

router.get('/get_user_status', async (ctx, next) => {
    console.log('getting user status')
    var id = ctx.request.query.id
    console.log('idid:', id)
    var collections = db.get('forms')
    var user = await collections.findOne({ _id: id })
    console.log(user)
    ctx.body = user
})

router.post('/post_test', async (ctx, next) => {
    const rb = ctx.request
    console.log('post test', rb.body)
    ctx.response.body = 'success'
})

router.post('/submit_form_post', async (ctx, next) => {
    const id = ctx.request.body.id
    const type = ctx.request.body.type
    const form = ctx.request.body.form
    var collections = db.get('forms')
    switch (type) {
        case "basic":
            const update_basic = await collections.update({
                _id: id
            }, {
                    $set: {
                        _basic_status: true,
                        _basic: JSON.parse(form)
                    }
                }
            )
            ctx.response.body = 666
            break
        case "job":
            const update_job = await collections.update({
                _id: id
            }, {
                    $set: {
                        _job_status: true,
                        _job: JSON.parse(form)
                    }
                })
            ctx.response.body = 666
            break
        case "honor":
            const update_honor = await collections.update({
                _id: id
            }, {
                    $set: {
                        _honor_status: true,
                        _honor: JSON.parse(form)
                    }
                })
            ctx.response.body = 666
            break
        case "satisfaction":
            const update_satisfaction = await collections.update({
                _id: id
            }, {
                    $set: {
                        _satisfaction_status: true,
                        _satisfaction: JSON.parse(form)
                    }
                })
            ctx.response.body = 666
            break
    }
})

//单位

router.get('/register', async (ctx, next) => {
    var user_name = ctx.query.user_name
    var user_password = ctx.query.user_password
    var user_id_card = ctx.query.user_id_card
    var collections = db.get('company_users')
    console.log('new user company')
    var user = await collections.findOne({ _user_name: user_name })
    if (user) {
        ctx.body = -1
        return
    }
    user = await collections.findOne({ _user_id_card: user_id_card })
    if (user) {
        ctx.body = -2
        return
    }
    var user_inserting = {
        _user_name: user_name,
        _user_password: user_password,
        _user_id_card: user_id_card
    }
    var result = await collections.insert(user_inserting)
    var user_id = result._id
    var form_collection = db.get('company_forms')
    var form_inserting = {
        _from_user: user_id.toString(),
        _confirmed: false,
        _basic: {
            '单位名称': null,
            '统一社会信用代码': null,
            '所属地域': null,
            '所属行业': null,
            '行业分类': null,
            '单位性质': null,
            '填报人': null,
            '联系电话': null,
            'QQ': null,
            '微信': null,
            '电子邮箱': null
        },
        _summary: [{
            "year": "2009",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2010",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2011",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2012",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2013",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2014",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2015",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2016",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2017",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }, {
            "year": "2018",
            "info": {
                "职工人数": null,
                "上一年度产值（万元）": null,
                "研发经费投入（万元）": null,
                "新产品销售收入（万元）": null,
                "专利申请授权数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "薪酬（月）": {
                    "3000元以下": null,
                    "3000-4000": null,
                    "4000-5000": null,
                    "5000-6000": null,
                    "6000-8000": null,
                    "8000-10000": null,
                    "10000-12000": null,
                    "12000-15000": null,
                    "15000-20000": null,
                    "20000元以上": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": null,
                    "num": null
                }]
            }
        }],
        _sum_in: [{
            "year": "2009",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2010",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2011",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2012",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2013",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2014",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2015",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2016",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2017",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2018",
            "info": {
                "流入人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }],
        _sum_out: [{
            "year": "2009",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2010",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2011",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2012",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2013",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2014",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2015",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2016",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2017",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }, {
            "year": "2018",
            "info": {
                "流出人数": null,
                "性别结构": {
                    "男": null,
                    "女": null
                },
                "年龄结构": {
                    "30岁以下": null,
                    "31-41": null,
                    "41-50": null,
                    "51-60": null,
                    "61岁以上": null
                },
                "学历结构": {
                    "博士研究生": null,
                    "硕士研究生": null,
                    "本科": null,
                    "大专": null,
                    "大专以下": null
                },
                "获得荣誉": {
                    "国家级": null,
                    "省部级": null,
                    "厅局级": null,
                    "其它": null,
                    "无": null
                },
                "岗位类别": [{
                    "cas": [],
                    "num": null
                }]
            }
        }],
        _out_status: [{
            "id": 0,
            "info": {
                "姓名": null,
                "身份证号": null,
                "年龄": null,
                '性别': null,
                '学历': null,
                "岗位类别": null,
                "从业年限（年）": null,
                "流入地域": null,
                '离职原因': null
            }
        }],
        _need: [{
            "id": 0,
            "info": {
                "需求岗位": null,
                "需求数量（人）": null,
                "年龄结构": null,
                '学历结构': null,
                '专业要求': null,
                "工作经验（年）": null,
                "职业资格证书": null,
                "岗位类别": null,
                '薪酬': null
            }
        }]
    }
    result = await form_collection.insert(form_inserting)
    ctx.body = 1
})

router.get('/login', async (ctx, next) => {
    var user_name = ctx.query.user_name
    var user_password = ctx.query.user_password
    ctx.body = await getUserId(user_name, user_password)
})

router.get('/changePassword', async (ctx, next) => {
    var user_name = ctx.query.user_name
    var user_id_card = ctx.query.user_id_card
    var user_new_pass = ctx.query.user_new_pass
    ctx.body = await changeUserPass(user_name, user_id_card, user_new_pass)
})

router.get('/companyFormGet', async (ctx, next) => {
    var user_name = ctx.query.user_name
    var user_password = ctx.query.user_password
    let id = await getUserId(user_name, user_password)
    let form = await getCompanyForm(id)
    ctx.body = form
})
//测试用的
router.get('/createTable', async (ctx, next) => {
    var user_name = ctx.query.user_name
    var user_password = ctx.query.user_password
    let id = await getUserId(user_name, user_password)
    let form = await getCompanyForm(id)
    let _basic = form._basic
    let _sum_in = await generateGangWeiLeiBie2(form._sum_in)
    let _sum_out = await generateGangWeiLeiBie2(form._sum_out)
    let _summary = await generateGangWeiLeiBie2(form._summary)
    let _out_status = form._out_status
    let _need = form._need
    console.log(_need[0]['info'])

    let data = [
        [_basic],
        _sum_in,
        _sum_out,
        _summary,
        _out_status,
        _need,
    ]
    await generateExcel(data)
    ctx.body = 'aaa'
})

router.get('/download/:name', async function (ctx) {
    var fileName = ctx.params.name+'.xlsx';
    console.log(fileName)
    // Set Content-Disposition to "attachment" to signal the client to prompt for download.
    // Optionally specify the filename of the download.
    // 设置实体头（表示消息体的附加信息的头字段）,提示浏览器以文件下载的方式打开
    // 也可以直接设置 ctx.set("Content-disposition", "attachment; filename=" + fileName);
    ctx.attachment(fileName);
    await send(ctx, fileName, { root: __dirname + '/excel' });
});

router.post('/companyFormSave', async (ctx) => {
    var user_name = ctx.query.user_name
    var user_password = ctx.query.user_password
    let data = ctx.request.body
    console.log(data)
    var collection = db.get('company_forms')
    const update_basic = await collection.update({
        _id: data._id
    }, data)
    //如果是提交状态，就生成对应的excel文件
    if(data._confirmed){
        //生成当前用户的Excel，以用户的id为名字
        let excelName = data._from_user
        console.log(excelName)
        let _basic = data._basic
        let _sum_in = await generateGangWeiLeiBie2(data._sum_in)
        let _sum_out = await generateGangWeiLeiBie2(data._sum_out)
        let _summary = await generateGangWeiLeiBie2(data._summary)
        let _out_status = data._out_status
        let _need = data._need
        let sourceData = [
            [_basic],
            _sum_in,
            _sum_out,
            _summary,
            _out_status,
            _need,
        ]
        await generateExcel(sourceData,excelName)
    }
    
    ctx.response.body = 'success'
})

app.use(cors({
    origin: function (ctx) {
        return "*"; // 允许来自所有域名请求
    },
    exposeHeaders: ['WWW-Authenticate', 'Server-Authorization'],
    maxAge: 5,
    credentials: true,
    allowMethods: ['GET', 'POST', 'DELETE'],
    allowHeaders: ['Content-Type', 'Authorization', 'Accept'],
}))

//跨域设置
app.use(async (ctx, next) => {
    ctx.set('Access-Control-Allow-Origin', '*');
    await next();
});
app.use(router.routes(), router.allowedMethods())
// SSL 证书配置
/*
let options = {
    key: fs.readFileSync("ssl/api.im-here.cn.key"),
    cert: fs.readFileSync("ssl/api.im-here.cn.pem")
}
*/

// 利用koa实例对象的callback()方法
// 结合http和https来启动服务器
http.createServer(app.callback()).listen(1234)
//https.createServer(options, app.callback()).listen(443)

module.exports = app
