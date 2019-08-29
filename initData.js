function initEveryYear(begin, end, data) {
    let dataArr = [];
    for (let i = begin; i <= end; i++) {
        dataArr.push({
            "year": i,
            "info": deepCopy(data)
        });
    }
    return dataArr;
}

function deepCopy(data) {
    if (typeof data !== 'object') {
        return data;
    }

    if (Array.isArray(data)) {
        let newArr = [];

        for (let item of data) {
            newArr.push(deepCopy(item));
        }
        return newArr;
    } else {
        let newObj = {};

        for (let attr in data) {
            newObj[attr] = deepCopy(data[attr]);
        }

        return newObj;
    }
}

const info_gender = {
    "男": null,
    "女": null
}

const info_age = {
    "25岁以下": null,
    "26-35岁": null,
    "36-45岁": null,
    "46-55岁": null,
    "56岁以上": null
}

const info_degree = {
    "博士研究生": null,
    "硕士研究生": null,
    "本科": null,
    "大专": null,
    "大专以下": null
}

const info_salary = {
    "3000以下": null,
    "3000-4000": null,
    "4000-5000": null,
    "5000-6000": null,
    "6000-8000": null,
    "8000-10000": null,
    "10000-12000": null,
    "12000-15000": null,
    "15000-20000": null,
    "20000以上": null
}

const info_people = [{
    "cas": [],
    "num": null
}]

const summary_info = {
    "职工人数": null,
    "性别结构": info_gender,
    "年龄结构": info_age,
    "学历结构": info_degree,
    "薪酬（元/月）": info_salary,
    "人员类别": info_people,
    "当年年度产值（万元）": null
}

const sum_in_info = {
    "流入人数": null,
    "性别结构": info_gender,
    "年龄结构": info_age,
    "学历结构": info_degree,
    "薪酬（元/月）": info_salary,
    "人员类别": info_people
}

const sum_out_info =  {
    "流出人数": null,
    "性别结构": info_gender,
    "年龄结构": info_age,
    "学历结构": info_degree,
    "薪酬（元/月）": info_salary,
    "人员类别": info_people,
    "离职原因（多选）": []
}

export const form_inserting = {
    _from_user: null,//user_id.toString(),
    _confirmed: false,
    _basic: {
        '所属地域': null,
        '所属行业': null,
        '单位名称（全称）': null,
        '统一社会信用代码/组织机构代码': null,
        '单位性质': null,
        '联系人': null,
        '联系电话': null,
        'QQ号': null,
        '微信号': null,
        '电子邮箱': null
    },
    _summary: initEveryYear(2009, 2018, summary_info),
    _summary_nei: initEveryYear(2009, 2018, summary_info),
    _summary_wai: initEveryYear(2009, 2018, summary_info),
    _sum_in: initEveryYear(2009, 2018, sum_in_info),
    _sum_in_nei: initEveryYear(2009, 2018, sum_in_info),
    _sum_in_wai: initEveryYear(2009, 2018, sum_in_info),
    _sum_out: initEveryYear(2009, 2018, sum_out_info),
    _sum_out_nei: initEveryYear(2009, 2018, sum_out_info),
    _sum_out_wai: initEveryYear(2009, 2018, sum_out_info),
    _out_status: [{
        "id": 0,
        "info": {
            "姓名": null,
            "身份证号": null,
            "年龄": null,
            '性别': null,
            "人员类别": null,
            "离职时间": null,
            "从业年限（年）": null,
            "流入地": null,
            '离职原因（多选）': null
        }
    }],
    _out_status_nei: [{
        "id": 0,
        "info": {
            "姓名": null,
            "身份证号": null,
            "年龄": null,
            '性别': null,
            "人员类别": null,
            "离职时间": null,
            "从业年限（年）": null,
            "流入地": null,
            '离职原因（多选）': null
        }
    }],
    _out_status_wai: [{
        "id": 0,
        "info": {
            "姓名": null,
            "身份证号": null,
            "年龄": null,
            '性别': null,
            "人员类别": null,
            "离职时间": null,
            "从业年限（年）": null,
            "流入地": null,
            '离职原因（多选）': null
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
            '薪酬（元/月）': null
        }
    }],
    _need_nei: [{
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
            '薪酬（元/月）': null
        }
    }],
    _need_wai: [{
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
            '薪酬（元/月）': null
        }
    }]
}

