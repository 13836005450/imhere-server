var koa = require('koa')
  , logger = require('koa-logger')
  , json = require('koa-json')
  , views = require('koa-views')
  , onerror = require('koa-onerror')
  , enforceHttps = require('koa-sslify').default
  , app = new koa()

var index = require('./routes/index')
var users = require('./routes/users')
var api = require('./routes/api')
var serve = require('koa-static')

// error handler
onerror(app)

// global middlewares
app.use(views('views', {
  root: __dirname + '/views',
  default: 'jade'
}))
app.use(require('koa-bodyparser')())
app.use(json())
app.use(logger())

app.use(function* (next) {
  var start = new Date
  yield next
  var ms = new Date - start
  console.log('%s %s - %s', this.method, this.url, ms)
})

app.use(serve('public/'));

app.use(enforceHttps())

// routes definitions
app.use(api.routes(), api.allowedMethods())

// error-handling
app.on('error', (err, ctx) => {
  console.error('server error', err, ctx)
})

module.exports = app
