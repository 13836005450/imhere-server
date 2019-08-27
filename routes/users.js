var router = require('koa-router')();

router.prefix('/users');

router.get('/', async function (ctx, next) {
  ctx.render('index', {
    title: 'Hello World foo!'
  });
})

module.exports = router;
