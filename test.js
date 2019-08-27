var crypto = require('crypto')
var md5 = crypto.createHash('md5')

var date = new Date()
console.log(date)
console.log(Number(date).toString())
var result = md5.update(Number(date).toString()).digest('hex')

console.log(result)