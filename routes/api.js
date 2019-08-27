var router = require('koa-router')()
var url = '127.0.0.1/test'
var db = require('monk')(url)
var crypto = require('crypto')

router.prefix('/api')

function checkDistance(a, b) {
    var lo1 = a.longitude
    var la1 = a.latitude
    var lo2 = b.longitude
    var la2 = b.latitude
    var La1 = la1 * Math.PI / 180.0
    var La2 = la2 * Math.PI / 180.0
    var La3 = La1 - La2;
    var Lb3 = lo1 * Math.PI / 180.0 - lo2 * Math.PI / 180.0
    var s = 2 * Math.asin(Math.sqrt(Math.pow(Math.sin(La3 / 2), 2) + Math.cos(La1) * Math.cos(La2) * Math.pow(Math.sin(Lb3 / 2), 2)))
    s = s * 6378.137
    s = Math.round(s * 10000) / 10
    s = s.toFixed(2)
    if (s <= 400) {
        return true
    } else {
        return false
    }
}

function locationCheck(cur_check_status, user_location) {
    if (cur_check_status === null) {
        return false
    }
    if (cur_check_status._gps_check === false) {
        return true
    } else {
        if (user_location.got === false) {
            return false
        }
        return checkDistance(cur_check_status._gps_location, user_location)
    }
}

function toBool(x) {
    if (x == 'true') {
        return true
    }
    if (x == 'false') {
        return false
    }
}

//ping!!!
router.get('/ping', async function (ctx, next) {
    ctx.body = 'pong!'
})

//new_user函数
router.get('/new_user', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const NAME = ctx.request.query.NAME
    const STUDENTID = ctx.request.query.STUDENTID
    //获取集合
    var user_collection = db.get('users')
    //检测用户是否存在
    var is_user_exists = await user_collection.count({ _openid: OPENID })
    var user_inserting = {
        _openid: OPENID,
        _name: NAME,
        _studentid: STUDENTID,
        _groups: []
    }
    if (!is_user_exists) {
        //插入用户
        const result = await user_collection.insert(user_inserting)
        console.log(OPENID, NAME, '注册成功')
        ctx.body = 1
    } else {
        console.log(OPENID, NAME, '已经注册过了')
        ctx.body = -1
    }
})

//group_create函数
router.get('/group_create', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GROUPNAME = ctx.request.query.GROUPNAME
    const GROUPPASS = ctx.request.query.GROUPPASS
    //获取集合
    var user_collection = db.get('users')
    var sysdata_collection = db.get('sysdata')
    var group_collection = db.get('groups')
    //当前房间号自增1
    const sysdata = await sysdata_collection.findOneAndUpdate({ name: 'groupcount' }, { $inc: { count: 1 } })
    //将创建好的小组加入到创建者的小组列表里
    const user = await user_collection.findOneAndUpdate({ _openid: OPENID }, { $push: { _groups: sysdata.count } })
    //在groups中创建新的小组
    var group_inserting = {
        _group_id: sysdata.count,
        _group_name: GROUPNAME,
        _group_pass: GROUPPASS,
        _creator_id: OPENID,
        _creator_name: user._name,
        _users: [],
        _history: [],
        _users_history: [],
        _last_check: '0'
    }
    const update_result_group = await group_collection.insert(group_inserting)
    ctx.body = GROUPNAME
})

//group_join
router.get('/group_join', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GROUPID = ctx.request.query.GROUPID
    const GROUPPASS = ctx.request.query.GROUPPASS
    //获取集合
    var user_collection = db.get('users')
    var group_collection = db.get('groups')
    //获取小组信息
    var group = await group_collection.findOne({ _group_id: Number(GROUPID) })
    if (group != null && GROUPPASS == group._group_pass) {
        //小组存在且密码正确
        var user = await user_collection.findOne({ _openid: OPENID })
        if (user._groups.indexOf(Number(GROUPID)) < 0) {
            //未加入这个小组
            //向用户的小组列表里插入此小组
            const update_result_user = await user_collection.update(
                { _openid: OPENID },
                { $addToSet: { _groups: Number(GROUPID) } }
            )
            var user_inserting = {
                _openid: OPENID,
                _name: user._name,
                _studentid: user._studentid
            }
            var user_history_inserting = {
                _openid: OPENID,
                _name: user._name,
                _studentid: user._studentid,
                _num_checked: 0,
                _num_unchecked: 0,
                _history: []
            }
            const update_result_group = await group_collection.update({
                _group_id: Number(GROUPID)
            }, {
                    $addToSet: {
                        _users: user_inserting,
                        _users_history: user_history_inserting
                    }
                }
            )
            ctx.body = 1
        } else {
            //已加入这个小组
            ctx.body = -2
        }
    } else {
        //小组不存在或密码不正确
        ctx.body = -1
    }
})

//get_homepage
router.get('/get_homepage', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GPSLOCATION = JSON.parse(ctx.request.query.GPSLOCATION)
    const CHECKKEY = ctx.request.query.CHECKKEY
    //获取集合
    var user_collection = db.get('users')
    var group_collection = db.get('groups')
    //获取用户信息
    var user = await user_collection.findOne({ _openid: OPENID })
    var user_groups = user._groups
    var homepage_data = []
    var cur_check_time = new Date()

    var checked = false

    for (let group_id of user_groups) {
        if (group_id == null)
            continue
        var group = await group_collection.findOne({ _group_id: Number(group_id) })
        if (group == null)
            continue
        var last_check_time = new Date(group._last_check)
        var users_history = group._users_history
        var user_history_cur = users_history.find(function (x) {
            return x._openid == OPENID
        })
        var cur_checked = false
        if (cur_check_time <= last_check_time && group._creator_id != OPENID && user_history_cur._history[user_history_cur._history.length - 1]._checked) {
            checked = true
        }
        const cur_check_status = group._history[group._history.length - 1]
        var checkcheckkey = group._history.length ? group._history[group._history.length - 1]._check_key : 'nullnull//nullnull'
        //console.log(checkcheckkey === CHECKKEY)
        if (cur_check_time <= last_check_time && group._creator_id != OPENID && locationCheck(cur_check_status, GPSLOCATION) && !(checkcheckkey === CHECKKEY) && !checked) {
            checked = true
            //点名成功
            if (!user_history_cur._history[user_history_cur._history.length - 1]._checked) {
                //找到user表并更改
                user_history_cur._history[user_history_cur._history.length - 1]._checked = true
                user_history_cur._num_checked++
                user_history_cur._num_unchecked--
                const update_result_group_user_history = await group_collection.update({
                    "_group_id": Number(group_id),
                    "_users_history._openid": OPENID
                }, {
                        $set: { "_users_history.$": user_history_cur }
                    })
                //找到group_history表并更改
                var group_history = group._history
                var group_history_user = group_history[group_history.length - 1]._user_unchecked.find(function (x) {
                    return x._openid == OPENID
                })
                const update_result_group_history = await group_collection.update({
                    "_group_id": Number(group_id),
                    "_history._id": group._history.length - 1
                }, {
                        $addToSet: { "_history.$._user_checked": group_history_user },
                        $pull: { "_history.$._user_unchecked": group_history_user },
                    })
            }
            cur_checked = user_history_cur._history[user_history_cur._history.length - 1]._checked
        }
        var homepage_data_ele = {
            group_id: group._group_id,
            group_name: group._group_name,
            group_creator_id: group._creator_id,
            group_creator_name: group._creator_name,
            group_permission: user._openid == group._creator_id ? "owner" : "user",
            last_check_time: group._last_check,
            num_unchecked: 0,
            cur_checked: cur_checked,
            check_key: group._history.length ? group._history[group._history.length - 1]._check_key : 'nullnull//nullnull'
        }
        if (homepage_data_ele.group_permission == "user") {
            var user_history = group._users_history.find(function (x) {
                return x._openid == OPENID
            })
            homepage_data_ele.num_unchecked += user_history._num_unchecked
        }
        homepage_data.push(homepage_data_ele)
    }
    ctx.body = homepage_data
})

//get_groupdata
router.get('/get_groupdata', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GROUPID = ctx.request.query.GROUPID
    //获取集合
    var group_collection = db.get('groups')
    //获取小组信息
    var group = await group_collection.findOne({ _group_id: Number(GROUPID) })
    ctx.body = group
})

//check_new
router.get('/check_new', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GROUPID = ctx.request.query.GROUPID
    const TIMEOFF = ctx.request.query.TIMEOFF
    const GPSCHECK = ctx.request.query.GPSCHECK
    const GPSLOCATION = ctx.request.query.GPSLOCATION
    //获取集合
    const group_collection = db.get('groups')
    //生成最后时间
    var last_check_time = new Date().getTime() + Number(TIMEOFF * 1000)
    //获取小组学生列表
    var group = await group_collection.findOne({ _group_id: Number(GROUPID) })
    var group_history = group._history
    var group_users = group._users
    var group_users_history = group._users_history
    for (let user_history of group_users_history) {
        let record = {
            _date: last_check_time,
            _checked: false
        }
        user_history._history.push(record)
        user_history._num_unchecked++
    }
    var md5 = crypto.createHash('md5')
    var curcheckkey = md5.update(Number(last_check_time).toString()).digest('hex')
    //构造本次点名的数据
    var check_data = {
        _id: group._history.length,
        _date: last_check_time,
        _gps_location: JSON.parse(GPSLOCATION),
        _gps_check: toBool(GPSCHECK),
        _user_checked: [],
        _user_unchecked: group_users,
        _check_key: curcheckkey
    }
    const update_result_group_user_history = await group_collection.update({
        "_group_id": Number(GROUPID)
    }, {
            $set: { "_users_history": group_users_history, "_last_check": last_check_time },
            $addToSet: { "_history": check_data }
        })
    ctx.body = group
})

//check_do
router.get('/check_do', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GROUPID = ctx.request.query.GROUPID
    const GPSLOCATION = JSON.parse(ctx.request.query.GPSLOCATION)
    //获取集合
    const group_collection = db.get('groups')
    //获取小组信息
    var group = await group_collection.findOne({ _group_id: Number(GROUPID) })
    //时间对比
    var cur_check_time = new Date()
    var last_check_time = new Date(group._last_check)
    const cur_check_status = group._history[group._history.length - 1]
    if (cur_check_time <= last_check_time) {
        //点名成功
        if (locationCheck(cur_check_status, GPSLOCATION)) {
            //位置判断
            var users_history = group._users_history
            var user_history_cur = users_history.find(function (x) {
                return x._openid == OPENID
            })
            if (!user_history_cur._history[user_history_cur._history.length - 1]._checked) {
                //找到user表并更改
                user_history_cur._history[user_history_cur._history.length - 1]._checked = true
                user_history_cur._num_checked++
                user_history_cur._num_unchecked--
                const update_result_group_user_history = await group_collection.update({
                    "_group_id": Number(GROUPID),
                    "_users_history._openid": OPENID
                }, {
                        $set: { "_users_history.$": user_history_cur }
                    })
                //找到group_history表并更改
                var group_history = group._history
                var group_history_user = group_history[group_history.length - 1]._user_unchecked.find(function (x) {
                    return x._openid == OPENID
                })
                const update_result_group_history = await group_collection.update({
                    "_group_id": Number(GROUPID),
                    "_history._id": group._history.length - 1
                }, {
                        $addToSet: { "_history.$._user_checked": group_history_user },
                        $pull: { "_history.$._user_unchecked": group_history_user },
                    })
                ctx.body = 1
            } else {
                ctx.body = 2
            }
        } else {
            ctx.body = -2
        }
    } else {
        ctx.body = -1
    }
})

//xzfs
router.get('/recheck', async function (ctx, next) {
    const GROUPID = ctx.request.query.GROUPID
    const TEACHEROPENID = ctx.request.query.OPENID
    const STUDENTOPENID = JSON.parse(ctx.request.query.STIDENTOPENID)
    //console.log("studentid",STUDENTOPENID)
    const DATE = Number(ctx.request.query.DATE)
    const group_collection = db.get('groups')
    const group = await group_collection.findOne({ _group_id: Number(GROUPID) })
    if (group._creator_id != TEACHEROPENID) {
        // console.log(group._creator_id,TEACHEROPENID)
        ctx.body = 0
    }
    else {
        // console.log(group)
        const group_historys = group._history.find(function (x) {
            return x._date == DATE
        })
        for (let i = 0; i < STUDENTOPENID.length; i++) {
            const user_info = group_historys._user_unchecked.find(function (x) {
                return x._openid == STUDENTOPENID[i]
            })
            if (user_info == null)
                break
            // console.log(user_info)
            //找到user_unchecked中的要补签的同学信息
            const update_result_group_history = await group_collection.findOneAndUpdate({
                _group_id: Number(GROUPID),
                "_history._date": DATE

            }, {
                    $pull: {
                        "_history.$._user_unchecked": user_info
                    },
                    $addToSet: {
                        "_history.$._user_checked": user_info
                    }
                })
            //console.log(update_result_group_history)
            let group_users_history_info = group._users_history.find(function (x) {
                return x._openid == STUDENTOPENID[i]
            })
            group_users_history_info._num_checked = group_users_history_info._num_checked + 1
            group_users_history_info._num_unchecked = group_users_history_info._num_unchecked - 1
            for (let j = 0; j < group_users_history_info._history.length; j++) {
                if (group_users_history_info._history[j]._date == Number(DATE)) {
                    group_users_history_info._history[j]._checked = true
                }
            }
            const update_result_group_user_history = await group_collection.update({
                _group_id: Number(GROUPID),
                "_users_history._openid": STUDENTOPENID[i]
            }, {
                    $set: {
                        "_users_history.$": group_users_history_info
                    }
                })
        }
        ctx.body = 1
    }
})

router.get('/update_user_info', async function (ctx, next) {
    const OPENID = ctx.request.query.OPENID
    const NAME = ctx.request.query.NAME.toString()
    const STUDENTID = ctx.request.query.STUDENTID
    const user_collection = db.get('users')
    const group_collection = db.get('groups')
    const user_info = await user_collection.findOne({ _openid: OPENID })
    const update_result_user = await user_collection.update({
        _openid: OPENID
    }, {
            $set: {
                _name: NAME,
                _studentid: STUDENTID
            }
        })
    for (let i = 0; i < user_info._groups.length; i++) {
        if (user_info._groups[i] == null)
            continue
        const group = await group_collection.findOne({ _group_id: user_info._groups[i] })
        if (group == null)
            continue
        if (group._creator_id == OPENID) {
            const group_creator_update = await group_collection.update({
                _group_id: user_info._groups[i]
            }, {
                    $set: {
                        _creator_name: NAME
                    }
                })
        } else {
            const group_users_update = await group_collection.findOneAndUpdate({
                _group_id: user_info._groups[i],
                "_users._openid": OPENID
            }, {
                    $set: {
                        "_users.$._name": NAME,
                        "_users.$._studentid": STUDENTID
                    }
                })
            const group_users_history_update = await group_collection.findOneAndUpdate({
                _group_id: user_info._groups[i],
                "_users_history._openid": OPENID
            }, {
                    $set: {
                        "_users_history.$._name": NAME,
                        "_users_history.$._studentid": STUDENTID
                    }
                })
        }
    }
    ctx.body = 1
})

router.get('/update_group_info', async function (ctx, next) {
    const OPENID = ctx.request.query.OPENID
    const GROUPID = ctx.request.query.GROUPID
    const PASSWORD = ctx.request.query.PASSWORD
    const NAME = ctx.request.query.NAME.toString()
    const group_collection = db.get('groups')
    const group = await group_collection.findOne({
        _group_id: Number(GROUPID)
    })
    if (group._creator_id != OPENID) {
        ctx.body = 0
    } else {
        const group_update = await group_collection.update({
            _group_id: Number(GROUPID)
        }, {
                $set: {
                    _group_name: NAME,
                    _group_pass: PASSWORD
                }
            })
        ctx.body = 1
    }
})

router.get('/group_user_exit', async function (ctx, next) {
    const OPENID = ctx.request.query.OPENID
    const GROUPID = Number(ctx.request.query.GROUPID)
    //获取小组id和用户id
    const user_collection = db.get('users')
    const group_collection = db.get('groups')

    var group_info = await group_collection.findOne({
        "_group_id": GROUPID
    })
    const user_info = group_info._users.find(function (x) {
        return x._openid == OPENID
    })
    const user_history_info = group_info._users_history.find(function (x) {
        return x._openid == OPENID
    })
    const update_result_group = await group_collection.update({
        "_group_id": GROUPID
    }, {
            $pull: {
                "_users": user_info,
                "_users_history": user_history_info
            }
        })
    var user = await user_collection.findOneAndUpdate({ _openid: OPENID }, {
        $pull: {
            _groups: GROUPID
        }
    })
    ctx.body = 1
})

router.get('/group_dismiss', async function (ctx, next) {
    //获取请求
    const OPENID = ctx.request.query.OPENID
    const GROUPID = ctx.request.query.GROUPID

    //获取集合
    var group_collection = db.get('groups')
    //获取小组信息
    var group = await group_collection.findOne({ _group_id: Number(GROUPID) })
    if (group._creator_id != OPENID) {
        ctx.body = -1
    } else {
        //小组存在,就解散小组
        const update_group = await group_collection.findOneAndDelete({ _group_id: Number(GROUPID) })
        ctx.body = 1
    }
})

module.exports = router