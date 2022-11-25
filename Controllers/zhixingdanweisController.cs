using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using yusuanxiangmu.DAL;
using yusuanxiangmu.Models;

namespace yusuanxiangmu.Controllers
{
    [Authorize]
    public class zhixingdanweisController : Controller
    {
        private caiwuContent db = new caiwuContent();

        // GET: zhixingdanweis
        public async Task<ActionResult> Index()
        {
            return View(await db.zhixingdanweis.ToListAsync());
        }
        
        // GET: zhixingdanweis/Create
        [Authorize(Roles = "Admin")]
        public ActionResult Create()
        {
            return View();
        }

        // POST: zhixingdanweis/Create
        // 为了防止“过多发布”攻击，请启用要绑定到的特定属性，有关 
        // 详细信息，请参阅 http://go.microsoft.com/fwlink/?LinkId=317598。
        [Authorize(Roles = "Admin")]
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "mingcheng")] zhixingdanwei zhixingdanwei)
        {
            if (ModelState.IsValid)
            {
                db.zhixingdanweis.Add(zhixingdanwei);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(zhixingdanwei);
        }        

        // GET: zhixingdanweis/Delete/5
        [Authorize(Roles = "Admin")]
        public async Task<ActionResult> Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            zhixingdanwei zhixingdanwei = await db.zhixingdanweis.FindAsync(id);
            if (zhixingdanwei == null)
            {
                return HttpNotFound();
            }
            return View(zhixingdanwei);
        }

        // POST: zhixingdanweis/Delete/5
        [Authorize(Roles = "Admin")]
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(string id)
        {
            zhixingdanwei zhixingdanwei = await db.zhixingdanweis.FindAsync(id);
            db.zhixingdanweis.Remove(zhixingdanwei);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
