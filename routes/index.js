var router = require('koa-router')();

router.get('/', async function (ctx, next) {
  ctx.body = 'pong!'
})

module.exports = router;
