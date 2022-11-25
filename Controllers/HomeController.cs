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
using System.IO;
using System.Data.OleDb;
using System.Transactions;
using System.Data.SqlClient;
using NPOI.XWPF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using Novacode;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace yusuanxiangmu.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private caiwuContent db = new caiwuContent();

        // GET: Home
        public async Task<ActionResult> Index(string sortOrder, string searchString, int? niandu, string shenhejieguo)
        {
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  select m;

            if (shenhejieguo == null || shenhejieguo == "全部")
            {

                ViewBag.shenhejieguozhi = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.shenhezhuangtai == shenhejieguo);
                ViewBag.shenhejieguozhi = shenhejieguo;
            }


            if (niandu.HasValue == true)
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == niandu);
                ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == jinnian);
                ViewBag.nianduzhi = jinnian;
            }

            //按照项目名称和下达日期进行升序或降序排列
            ViewBag.xiangmuleibie = String.IsNullOrEmpty(sortOrder) ? "Name_desc" : "";
            ViewBag.xiadariqi = sortOrder == "Date" ? "Date_desc" : "Date";
            //var students = from s in db.xiangmuzongbiaos
            //               select s;

            //模糊查询
            if (!String.IsNullOrEmpty(searchString))
            {
                //xiangmuzongbiao = xiangmuzongbiao.Where(s => s.mingcheng.ToUpper().Contains(searchString.ToUpper())
                //    || s.xiangmuID.ToUpper().Contains(searchString.ToUpper()));

                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.mingcheng.ToUpper().Contains(searchString.ToUpper())
                    || s.xiangmuleibie.ToUpper().Contains(searchString.ToUpper()));//                
            }

            ViewBag.searchzhi = searchString;

            ViewBag.sortOrderzhi = sortOrder;

            switch (sortOrder)
            {
                case "Name_desc":
                    xiangmuzongbiao = xiangmuzongbiao.OrderByDescending(s => s.xiangmuleibie);
                    break;
                case "Date":
                    xiangmuzongbiao = xiangmuzongbiao.OrderBy(s => s.xiadariqi);
                    break;
                case "Date_desc":
                    xiangmuzongbiao = xiangmuzongbiao.OrderByDescending(s => s.xiadariqi);
                    break;
                default:
                    xiangmuzongbiao = xiangmuzongbiao.OrderBy(s => s.xiangmuleibie);
                    break;
            }
            return View(await xiangmuzongbiao.ToListAsync());
        }


        // GET: Home/Details/5  带有审核功能的实现
        [Authorize]
        public async Task<ActionResult> Details(int id, string sortOrder, string searchstring, string shenhejieguo, int? niandu)
        {
            ViewBag.nianduzhi = niandu;
            ViewBag.searchzhi = searchstring;
            ViewBag.shenhejieguozhi = shenhejieguo;
            ViewBag.sortOrderzhi = sortOrder;
            xiangmuzongbiao xiangmuzongbiao = await db.xiangmuzongbiaos.FindAsync(id);
            if (xiangmuzongbiao == null)
            {
                return HttpNotFound();
            }
            return View(xiangmuzongbiao);
        }


        // POST: Home/Edit/5  带有审核功能的实现
        // 为了防止“过多发布”攻击，请启用要绑定到的特定属性，有关 
        // 详细信息，请参阅 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [Authorize]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Details([Bind(Include = "ID,niandu,xiangmuleibie,mingcheng,shangcaixiangmuhao,lierulanmu,bianhao,zhixingdanwei,lianxiren,jine,xiangmufenlei,zijinlaiyuan,zijinxingzhi,xiabobumen,xiadariqi,xiadaqingkuang,xiaozhanghui,dangweihui,lururen,lururiqi,shenheren,shenheriqi,beizhu,shenpiliucheng")] xiangmuzongbiao xiangmuzongbiao)
        {
            if (ModelState.IsValid)
            {
                db.Entry(xiangmuzongbiao).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");                //这个是否要改成长链接 待定
            }
            return View(xiangmuzongbiao);
        }



        // GET: Home/Create
        [Authorize(Roles = "工作人员,科长")]
        public ActionResult Create()
        {
            var GenreLst = new List<string>();

            var GenreQry = from d in db.zhixingdanweis
                           orderby d.mingcheng
                           select d.mingcheng;

            GenreLst.AddRange(GenreQry.Distinct());
            //ViewBag.zxdanwei = new SelectList(GenreLst);
            ViewBag.zxdanwei = GenreLst;

            return View();
        }


        // POST: Home/Create
        // 为了防止“过多发布”攻击，请启用要绑定到的特定属性，有关 
        // 详细信息，请参阅 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [Authorize(Roles = "工作人员,科长")]
        [ValidateAntiForgeryToken]
        //   public async Task<ActionResult> Create([Bind(Include = "xiangmuID,mingcheng,xiangmudalei,zijinlaiyuan,lierulanmu,wenjianbianhao,zhixingdanwei,lianxiren,jine,xiadariqi,zijinxingzhi,xiabobumen,lururen,shenheren")] xiangmuzongbiao xiangmuzongbiao)
        public ActionResult Create([Bind(Include = "ID,niandu,xiangmuleibie,mingcheng,shangcaixiangmuhao,lierulanmu,bianhao,zhixingdanwei,lianxiren,jine,xiangmufenlei,zijinlaiyuan,xiangmulaiyuan,zijinxingzhi,xiabobumen,xiadariqi,xiadaqingkuang,xiaozhanghui,dangweihui,beizhu,shenpiliucheng")] xiangmuzongbiao xiangmuzongbiao)
        //这里漏掉了一个资金归属                                                                                          这里加里一个年度
        {
            xiangmuzongbiao.shenhezhuangtai = "未审核";
            xiangmuzongbiao.lururen = User.Identity.Name.ToString();
            xiangmuzongbiao.lururiqi = DateTime.Now.ToString("yyyy-MM-dd");
            xiangmuzongbiao.shenheren = "";
            xiangmuzongbiao.shenheriqi = "";
            try
            {
                if (ModelState.IsValid)
                {
                    db.xiangmuzongbiaos.Add(xiangmuzongbiao);
                    // await db.SaveChangesAsync();
                    db.SaveChanges();
                    return RedirectToAction("Index");
                }
            }
            //catch (Exception e)
            catch (DataException)
            {
                // ModelState.AddModelError("", e.ToString());
                ModelState.AddModelError("", "不能保存，有可能是项目号或者是项目名称重复等！");
            }
            var GenreLst = new List<string>();
            var GenreQry = from d in db.zhixingdanweis
                           orderby d.mingcheng
                           select d.mingcheng;
            GenreLst.AddRange(GenreQry.Distinct());
            //ViewBag.zxdanwei = new SelectList(GenreLst);
            ViewBag.zxdanwei = GenreLst;

            return View();    //返回错误描述
        }

        // GET: Home/Edit/5
        [Authorize(Roles = "工作人员")]
        public async Task<ActionResult> Edit(int id, string sortOrder, string searchstring, string shenhejieguo, int? niandu)
        {
            ViewBag.nianduzhi = niandu;
            ViewBag.searchzhi = searchstring;
            ViewBag.shenhejieguozhi = shenhejieguo;
            ViewBag.sortOrderzhi = sortOrder;

            xiangmuzongbiao xiangmuzongbiao = await db.xiangmuzongbiaos.FindAsync(id);
            if (xiangmuzongbiao == null)
            {
                return HttpNotFound();
            }
            //这里是为执行单位加上去的
            var GenreLst = new List<string>();

            var GenreQry = from d in db.zhixingdanweis
                           orderby d.mingcheng
                           select d.mingcheng;
            GenreLst.AddRange(GenreQry.Distinct());
            //ViewBag.zxdanwei = new SelectList(GenreLst);
            ViewBag.zxdanwei = GenreLst;
            return View(xiangmuzongbiao);
        }

        // POST: Home/Edit/5
        // 为了防止“过多发布”攻击，请启用要绑定到的特定属性，有关 
        // 详细信息，请参阅 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [Authorize(Roles = "工作人员")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit(string nianduzhi, string searchzhi, string shenhejieguozhi, string sortOrderzhi, [Bind(Include = "ID,niandu,xiangmuleibie,mingcheng,shangcaixiangmuhao,lierulanmu,bianhao,zhixingdanwei,lianxiren,jine,xiangmufenlei,zijinlaiyuan,xiangmulaiyuan,zijinxingzhi,xiabobumen,xiadariqi,xiadaqingkuang,lururen,lururiqi,shenheren,shenheriqi,shenhezhuangtai,xiaozhanghui,dangweihui,beizhu,shenpiliucheng")] xiangmuzongbiao xiangmuzongbiao)//,niandu,xiaozhanghui,dangweihui
        {
            if (ModelState.IsValid)
            {
                db.Entry(xiangmuzongbiao).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index", "Home", new { sortOrder = sortOrderzhi, searchString = searchzhi, niandu = nianduzhi, shenhejieguo = shenhejieguozhi });
            }

            //这里是为执行单位加上去的
            var GenreLst = new List<string>();

            var GenreQry = from d in db.zhixingdanweis
                           orderby d.mingcheng
                           select d.mingcheng;
            GenreLst.AddRange(GenreQry.Distinct());
            //ViewBag.zxdanwei = new SelectList(GenreLst);
            ViewBag.zxdanwei = GenreLst;
            return View(xiangmuzongbiao);
        }

        // GET: Home/shenhe/5
        [Authorize(Roles = "科长")]
        public async Task<ActionResult> shenhe(int? id, string sortOrder, string searchString, string shenhejieguo, int? niandu)
        {
            ViewBag.nianduzhi = niandu;
            ViewBag.searchzhi = searchString;
            ViewBag.shenhejieguozhi = shenhejieguo;
            ViewBag.sortOrderzhi = sortOrder;

            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            xiangmuzongbiao xiangmuzongbiao = await db.xiangmuzongbiaos.FindAsync(id);
            if (xiangmuzongbiao == null)
            {
                return HttpNotFound();
            }

            return View(xiangmuzongbiao);
        }

        // POST: Home/shenhe/5
        // 为了防止“过多发布”攻击，请启用要绑定到的特定属性，有关 
        // 详细信息，请参阅 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [Authorize(Roles = "科长")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> shenhe(string nianduzhi, string searchzhi, string shenhejieguozhi, string sortOrderzhi, xiangmuzongbiao xiangmuzongbiao)
        {
            xiangmuzongbiao.shenheren = User.Identity.Name.ToString();
            xiangmuzongbiao.shenheriqi = DateTime.Now.ToString("yyyy-MM-dd");
            if (ModelState.IsValid)
            {
                db.Entry(xiangmuzongbiao).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index", "Home", new { sortOrder = sortOrderzhi, searchString = searchzhi, niandu = nianduzhi, shenhejieguo = shenhejieguozhi });
            }
            return View(xiangmuzongbiao);
        }


        // GET: Home/shenhe/5
        [Authorize(Roles = "处长")]
        public async Task<ActionResult> chehui(int? id, string sortOrder, string searchString, string shenhejieguo, int? niandu)     //数值型作为参数传递时不可为空，可在 int 后面加 “？”允许参数可空
        {
            ViewBag.nianduzhi = niandu;
            ViewBag.searchzhi = searchString;
            ViewBag.shenhejieguozhi = shenhejieguo;
            ViewBag.sortOrderzhi = sortOrder;
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            xiangmuzongbiao xiangmuzongbiao = await db.xiangmuzongbiaos.FindAsync(id);
            if (xiangmuzongbiao == null)
            {
                return HttpNotFound();
            }
            return View(xiangmuzongbiao);
        }

        // POST: Home/shenhe/5
        // 为了防止“过多发布”攻击，请启用要绑定到的特定属性，有关 
        // 详细信息，请参阅 http://go.microsoft.com/fwlink/?LinkId=317598。
        [HttpPost]
        [Authorize(Roles = "处长")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> chehui(string nianduzhi, string searchzhi, string shenhejieguozhi, string sortOrderzhi, [Bind(Include = "ID,niandu,xiangmuleibie,mingcheng,shangcaixiangmuhao,lierulanmu,bianhao,zhixingdanwei,lianxiren,jine,xiangmufenlei,zijinlaiyuan,xiangmulaiyuan,zijinxingzhi,xiabobumen,xiadariqi,xiadaqingkuang,lururen,lururiqi,shenheren,shenheriqi,shenhezhuangtai,xiaozhanghui,dangweihui,beizhu")] xiangmuzongbiao xiangmuzongbiao)
        {
            xiangmuzongbiao.shenheren = User.Identity.Name.ToString();
            xiangmuzongbiao.shenheriqi = DateTime.Now.ToString("yyyy-MM-dd");
            xiangmuzongbiao.shenhezhuangtai = "未通过";
            if (ModelState.IsValid)
            {
                db.Entry(xiangmuzongbiao).State = EntityState.Modified;
                await db.SaveChangesAsync();
                //return RedirectToAction("Index");

                return RedirectToAction("Index", "Home", new { sortOrder = sortOrderzhi, searchString = searchzhi, niandu = nianduzhi, shenhejieguo = shenhejieguozhi });
            }
            return View(xiangmuzongbiao);
        }

        // GET: Home/Delete/5
        [Authorize(Roles = "处长,工作人员")]
        public async Task<ActionResult> Delete(int? id, string shenhezhuangtai, string sortOrder, string searchString, string shenhejieguo, int? niandu)
        {
            ViewBag.nianduzhi = niandu;
            ViewBag.searchzhi = searchString;
            ViewBag.shenhejieguozhi = shenhejieguo;
            ViewBag.sortOrderzhi = sortOrder;
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            xiangmuzongbiao xiangmuzongbiao = await db.xiangmuzongbiaos.FindAsync(id);
            if (User.IsInRole("处长") == true)
            {
                if (xiangmuzongbiao == null)
                {
                    return HttpNotFound();
                }
                return View(xiangmuzongbiao);
            }

            if (shenhezhuangtai != "通过" && shenhezhuangtai != null)
            {
                if (xiangmuzongbiao == null)
                {
                    return HttpNotFound();
                }
                return View(xiangmuzongbiao);
            }
            return RedirectToAction("Index");
        }

        // POST: Home/Delete/5
        [HttpPost, ActionName("Delete")]
        [Authorize(Roles = "处长,工作人员")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id, string nianduzhi, string searchzhi, string shenhejieguozhi, string sortOrderzhi)
        {
            xiangmuzongbiao xiangmuzongbiao = await db.xiangmuzongbiaos.FindAsync(id);
            db.xiangmuzongbiaos.Remove(xiangmuzongbiao);
            await db.SaveChangesAsync();
            return RedirectToAction("Index", "Home", new { sortOrder = sortOrderzhi, searchString = searchzhi, niandu = nianduzhi, shenhejieguo = shenhejieguozhi });
        }

        //GET: Home
        public async Task<ActionResult> zhichuhuizong(string xmleibie, int? niandu)             //这里的支出汇总已经照着index方法进行更改，并且把xiangmuzongbiaos改为另外xianngmuzongbiao        
        {
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  where m.shenhezhuangtai == "通过"
                                  select m;

            if (niandu.HasValue == true)
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == niandu);
                ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == jinnian);
                ViewBag.nianduzhi = jinnian;
            }


            //ViewBag.xingxi = "支出汇总表";
            //按名称搜索
            //var xiangmuzongbiaos = from m in db.xiangmuzongbiaos
            //                       select m;
            if (String.IsNullOrEmpty(xmleibie) || xmleibie == "全部")
            {
                ViewBag.leibie = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.xiangmuleibie == xmleibie);
                ViewBag.leibie = xmleibie;
            }

            var zhichu = new List<zhichuzongbiao>();     //这里是新加的

            //压缩传送表单数据量
            var zhichutemp = from d in xiangmuzongbiao
                             orderby d.xiadariqi       //项目分类 ，资金来源
                             select new zhichuzongbiao
                             {
                                 xiangmufenlei = d.xiangmufenlei,
                                 mingcheng = d.mingcheng,
                                 shangcaixiangmuhao = d.shangcaixiangmuhao,
                                 lierulanmu = d.lierulanmu,
                                 xiangmulaiyuan = d.xiangmulaiyuan,
                                 zhixingdanwei = d.zhixingdanwei,
                                 lianxiren = d.lianxiren,
                                 jine = d.jine,
                                 niandu = d.niandu,
                                 xiangmuleibie = d.xiangmuleibie,
                                 zijinlaiyuan = d.zijinlaiyuan
                             };
            zhichu.AddRange(zhichutemp.Distinct());
            //return View(zhichutemp.ToList());
            return View(await zhichutemp.ToListAsync());
        }


        public async Task<ActionResult> shouruhuizong(string xmleibie, int? niandu)             //这里的支出汇总已经照着index方法进行更改，并且把xiangmuzongbiaos改为另外xianngmuzongbiao        
        {
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  select m;

            if (niandu.HasValue == true)
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == niandu && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == jinnian && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = jinnian;
            }

            if (String.IsNullOrEmpty(xmleibie) || xmleibie == "全部")
            {
                ViewBag.leibie = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.xiangmuleibie == xmleibie);
                ViewBag.leibie = xmleibie;
            }

            var shouru = new List<shouruzongbiao>();  //新加入的

            var shourutemp = from d in xiangmuzongbiao
                             orderby d.xiadariqi

                             select new shouruzongbiao     //控制xiadariqi正常显示的是shouruzongbiao，不是shourubiaodaochu
                             {
                                 mingcheng = d.mingcheng,
                                 xiangmulaiyuan = d.xiangmulaiyuan,    //这里到底是“文件编号”还是“项目ID”有待确定
                                 xiangmuleibie = d.xiangmuleibie,
                                 zijinlaiyuan = d.zijinlaiyuan,

                                 zijinxingzhi = d.zijinxingzhi,
                                 zhixingdanwei = d.zhixingdanwei,
                                 xiabobumen = d.xiabobumen,
                                 xiadariqi = d.xiadariqi,
                                 jine = d.jine,
                                 niandu = d.niandu
                             };

            shouru.AddRange(shourutemp.Distinct());
            return View(await shourutemp.ToListAsync());
        }

        //NPOI导出电子表格
        public FileResult Exportzhichu(string searchString, int? niandu)//
        {
            //string schoolname = "401";
            //创建Excel文件的对象
            NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //添加一个sheet
            NPOI.SS.UserModel.ISheet sheet1 = book.CreateSheet("Sheet1");
            //获取list数据
            //List<TB_STUDENTINFOModel> listRainInfo = m_BLL.GetSchoolListAATQ(schoolname);            

            var xiangmuzongbiaos = from m in db.xiangmuzongbiaos
                                   where m.shenhezhuangtai == "通过"
                                   select m;
            if (!String.IsNullOrEmpty(searchString))
            {
                xiangmuzongbiaos = xiangmuzongbiaos.Where(s => s.mingcheng.ToUpper().Contains(searchString.ToUpper())
                    || s.shangcaixiangmuhao.ToUpper().Contains(searchString.ToUpper()));//&& s.niandu == niandu)
            }

            if (niandu.HasValue == true)
            {
                xiangmuzongbiaos = xiangmuzongbiaos.Where(s => s.niandu == niandu);
                //ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiaos = xiangmuzongbiaos.Where(s => s.niandu == jinnian);
                //ViewBag.nianduzhi = jinnian;
            }

            NPOI.SS.UserModel.IRow row_biaoti = sheet1.CreateRow(0);

            ICell cell1 = row_biaoti.CreateCell(0);
            ICellStyle style = book.CreateCellStyle();
            //设置单元格的样式：水平对齐居中
            style.Alignment = HorizontalAlignment.Center;
            //新建一个字体样式对象
            IFont font = book.CreateFont();
            //设置字体加粗样式
            font.Boldweight = short.MaxValue;
            font.FontHeightInPoints = 20;                            //设置字体大小
            //使用SetFont方法将字体样式添加到单元格样式中 
            style.SetFont(font);
            //将新的样式赋给单元格
            cell1.CellStyle = style;                                          //在指定的cell中使用样式

            ICellStyle cellstyle_jine = book.CreateCellStyle();
            cellstyle_jine.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

            string niandu_biaoti = niandu.ToString();                        //设置标题年度      
            cell1.SetCellValue(niandu_biaoti + "年支出汇总表");
            //cell1.SetCellValue(niandu_biaoti);

            row_biaoti.Height = 30 * 20;                                      //设置行高
            sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 7));         //合并单元格

            NPOI.SS.UserModel.IRow row_zhibiao = sheet1.CreateRow(1);

            sheet1.AddMergedRegion(new CellRangeAddress(1, 1, 0, 7));            //(第一行，最后一行，第一列，最后一列)

            ICell cell_zhibiaoshijian = row_zhibiao.CreateCell(0);            //数字有待研究
            ICellStyle style1 = book.CreateCellStyle();                       //设置另外一种样式
            style1.Alignment = HorizontalAlignment.Right;                     //样式的内容为右对齐
            cell_zhibiaoshijian.CellStyle = style1;                           //使用样式                

            var zhibiao = "制表时间：" + DateTime.Now.ToShortDateString();
            //row_zhibiao.CreateCell(0).SetCellValue(zhibiao);                   //这里引用样式有问题
            cell_zhibiaoshijian.SetCellValue(zhibiao);

            //给sheet1添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(2);
            // row1.CreateCell(0).SetCellValue("项目分类");                 //这是新加上的
            row1.CreateCell(0).SetCellValue("序号");

            row1.CreateCell(1).SetCellValue("项目号");
            row1.CreateCell(2).SetCellValue("项目名称");

            row1.CreateCell(3).SetCellValue("列入栏目");            //这里要更改相对应的序号

            row1.CreateCell(4).SetCellValue("文件编号");
            row1.CreateCell(5).SetCellValue("执行单位");
            row1.CreateCell(6).SetCellValue("联系人");
            row1.CreateCell(7).SetCellValue("项目金额(万元)");

            //设置列宽
            sheet1.SetColumnWidth(0, 10 * 256);   // 在EXCEL文档中实际列宽为49.29
            sheet1.SetColumnWidth(1, 30 * 256);
            sheet1.SetColumnWidth(2, 30 * 256);
            sheet1.SetColumnWidth(3, 30 * 256);
            sheet1.SetColumnWidth(4, 30 * 256);
            sheet1.SetColumnWidth(5, 20 * 256);
            sheet1.SetColumnWidth(6, 20 * 256);
            sheet1.SetColumnWidth(7, 20 * 256);

            string[] fenlei = { "教学", "科研", "行政", "后勤", "离退休" };
            string[] laiyuan = { "网上指标", "基本户", "校内追加" };
            string[] fenleixuhao = { "一、", "二、", "三、", "四、", "五、" };
            string[] laiyuanxuhao = { "1", "2", "3" };

            NPOI.SS.UserModel.IRow row_zongji = sheet1.CreateRow(3);
            row_zongji.CreateCell(1).SetCellValue("总计");
            decimal zongji_sum = 0;

            int k = 4;//3                 

            for (int i = 0; i < 5; i++)
            {
                int fenlei_weizhi = k;
                decimal fenlei_sum = 0;

                for (int j = 0; j < 3; j++)
                {
                    k = k + 1;
                    row1 = sheet1.CreateRow(k);                   // 原先是:row1 = sheet1.CreateRow(k+1);
                    row1.CreateCell(0).SetCellValue(laiyuanxuhao[j]);
                    row1.CreateCell(1).SetCellValue(laiyuan[j]);

                    var GenreLst = new List<zhichuzongbiao>();
                    var GenreQry = from d in xiangmuzongbiaos
                                   orderby d.xiadariqi
                                   select new zhichuzongbiao
                                   {
                                       shangcaixiangmuhao = d.shangcaixiangmuhao,
                                       mingcheng = d.mingcheng,

                                       lierulanmu = d.lierulanmu,
                                       xiangmulaiyuan = d.xiangmulaiyuan,
                                       zhixingdanwei = d.zhixingdanwei,
                                       lianxiren = d.lianxiren,
                                       jine = d.jine,

                                       xiangmufenlei = d.xiangmuleibie,                //这是新加上的        这里有改动  待定 0311

                                       zijinlaiyuan = d.zijinlaiyuan,

                                       niandu = d.niandu,                             //这里加上一个年度，在 标题上的 年度数值直接使用这里的值
                                   };
                    GenreLst.AddRange(GenreQry);     //GenreQry.Distinct():去除相同的元素
                    GenreLst = GenreLst.Where(d => d.xiangmufenlei == fenlei[i] && d.zijinlaiyuan == laiyuan[j]).ToList();
                    int GenreLst_count = GenreLst.Count();
                    // GenreLst.Where(d => d.xiangmufenlei == fenlei[i] && d.zijinlaiyuan == laiyuan[j]);                                               

                    if (GenreLst_count == 0)
                    {
                        row1.CreateCell(7).SetCellValue(0);
                        // k = k + 1;
                    }
                    else
                    {
                        ICell cell_fenleixiaoji = row1.CreateCell(7);
                        cell_fenleixiaoji.SetCellFormula("sum(h" + (k + 2).ToString() + ":" + "h" + (k + 1 + GenreLst_count).ToString() + ")");
                        cell_fenleixiaoji.CellStyle = cellstyle_jine;
                    }

                    int h = 1;

                    foreach (var item in GenreLst)
                    //foreach (var item in zhichu)
                    {
                        row1 = sheet1.CreateRow(k + 1);
                        ICell cell = row1.CreateCell(0);
                        cell.SetCellValue(laiyuanxuhao[j] + "." + h.ToString());

                        cell = row1.CreateCell(1);
                        cell.SetCellValue(item.shangcaixiangmuhao);

                        cell = row1.CreateCell(2);
                        cell.SetCellValue(item.mingcheng);

                        cell = row1.CreateCell(3);
                        cell.SetCellValue(item.lierulanmu);

                        cell = row1.CreateCell(4);
                        cell.SetCellValue(item.xiangmulaiyuan);

                        cell = row1.CreateCell(5);
                        cell.SetCellValue(item.zhixingdanwei);

                        cell = row1.CreateCell(6);
                        cell.SetCellValue(item.lianxiren);

                        fenlei_sum = fenlei_sum + item.jine;
                        float s1 = float.Parse(item.jine.ToString());
                        cell = row1.CreateCell(7);
                        cell.SetCellValue(s1);

                        //ICellStyle cellstyle = book.CreateCellStyle();
                        //cellstyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                        cell.CellStyle = cellstyle_jine;

                        h = h + 1;
                        k = k + 1;
                    }
                }

                k = k + 1;

                row1 = sheet1.CreateRow(fenlei_weizhi);
                row1.CreateCell(0).SetCellValue(fenleixuhao[i]);
                row1.CreateCell(1).SetCellValue(fenlei[i]);
                ICell feileisum_cell = row1.CreateCell(7);
                feileisum_cell.SetCellValue(fenlei_sum.ToString());//分类总计有更改


                feileisum_cell.CellStyle = style1;

                zongji_sum = zongji_sum + fenlei_sum;
            }
            row_zongji.CreateCell(7).SetCellValue(zongji_sum.ToString());
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return File(ms, "application/vnd.ms-excel", "支出总表.xls");
        }

        public FileResult Exportshouru(int? niandu)
        {
            //string schoolname = "401";
            //创建Excel文件的对象
            NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
            //添加一个sheet
            NPOI.SS.UserModel.ISheet sheet1 = book.CreateSheet("Sheet1");
            //获取list数据
            //List<TB_STUDENTINFOModel> listRainInfo = m_BLL.GetSchoolListAATQ(schoolname);

            var xiangmuzongbiaos = from m in db.xiangmuzongbiaos
                                   orderby m.xiadariqi
                                   where m.shenhezhuangtai == "通过"
                                   select m;            

            if (niandu.HasValue == true)
            {
                xiangmuzongbiaos = xiangmuzongbiaos.Where(s => s.niandu == niandu);
                //ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiaos = xiangmuzongbiaos.Where(s => s.niandu == jinnian);
                //ViewBag.nianduzhi = jinnian;
            }

            //表头
            NPOI.SS.UserModel.IRow row_biaoti = sheet1.CreateRow(0);
            ICell cell1 = row_biaoti.CreateCell(0);
            ICellStyle style = book.CreateCellStyle();
            //设置单元格的样式：水平对齐居中
            style.Alignment = HorizontalAlignment.Center;
            //新建一个字体样式对象
            IFont font = book.CreateFont();
            //设置字体加粗样式
            font.Boldweight = short.MaxValue;
            font.FontHeightInPoints = 20;                            //设置字体大小
                                                                     //使用SetFont方法将字体样式添加到单元格样式中 
            style.SetFont(font);
            //将新的样式赋给单元格
            cell1.CellStyle = style;
            //在指定的cell中使用样式

            string niandu_biaoti = niandu.ToString();                        //设置标题年度      

            cell1.SetCellValue(niandu_biaoti + "年收入汇总表");

            row_biaoti.Height = 30 * 20;                                      //设置行高
            sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 9));         //合并单元格

            ICellStyle cellstyle_jine = book.CreateCellStyle();                          //总的设置金额小数点数目
            cellstyle_jine.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");

            sheet1.AddMergedRegion(new CellRangeAddress(1, 1, 0, 9));            //(第一行，最后一行，第一列，最后一列)

            //制表时间
            NPOI.SS.UserModel.IRow row_zhibiao = sheet1.CreateRow(1);          //括号中的1 为 第一行
            ICell cell_zhibiaoshijian = row_zhibiao.CreateCell(0);            //数字有待研究
            ICellStyle style1 = book.CreateCellStyle();                       //设置另外一种样式
            style1.Alignment = HorizontalAlignment.Right;                     //样式的内容为右对齐
            cell_zhibiaoshijian.CellStyle = style1;                           //使用样式                

            var zhibiao = "制表时间：" + DateTime.Now.ToShortDateString();
            //row_zhibiao.CreateCell(0).SetCellValue(zhibiao);                  //这里引用样式有问题,应该在行中的单元格引用样式，不是单元行中引用样式
            cell_zhibiaoshijian.SetCellValue(zhibiao);

            //给sheet1添加第一行的头部标题
            NPOI.SS.UserModel.IRow row1 = sheet1.CreateRow(2);
            row1.CreateCell(0).SetCellValue("序号");

            row1.CreateCell(1).SetCellValue("项目名称");
            //这里少一个项目号

            row1.CreateCell(2).SetCellValue("项目分类");
            row1.CreateCell(3).SetCellValue("资金来源");
            //这里少一个项目来源

            row1.CreateCell(4).SetCellValue("资金性质");

            row1.CreateCell(5).SetCellValue("执行单位");
            row1.CreateCell(6).SetCellValue("下拨部门");
            row1.CreateCell(7).SetCellValue("下达日期");
            row1.CreateCell(8).SetCellValue("编号");          //文件编号  在这里更改为编号；  少一个项目号，少一个资金来源
            row1.CreateCell(9).SetCellValue("项目金额(万元)");

            //设置列宽
            sheet1.SetColumnWidth(0, 10 * 256);   // 在EXCEL文档中实际列宽为49.29
            sheet1.SetColumnWidth(1, 30 * 256);
            sheet1.SetColumnWidth(2, 18 * 256);
            sheet1.SetColumnWidth(3, 18 * 256);
            sheet1.SetColumnWidth(4, 18 * 256);
            sheet1.SetColumnWidth(5, 18 * 256);
            sheet1.SetColumnWidth(6, 18 * 256);
            sheet1.SetColumnWidth(7, 20 * 256);
            sheet1.SetColumnWidth(8, 30 * 256);
            sheet1.SetColumnWidth(9, 20 * 256);

            string[] fenlei = { "教学", "科研", "行政", "后勤", "离退休" };
            string[] laiyuan = { "网上指标", "基本户", "校内追加" };
            string[] fenleixuhao = { "一、", "二、", "三、", "四、", "五、" };
            string[] laiyuanxuhao = { "1", "2", "3" };

            NPOI.SS.UserModel.IRow row_zongji = sheet1.CreateRow(3);
            row_zongji.CreateCell(1).SetCellValue("总计");
            decimal zongji_sum = 0;

            int k = 4;//3

            for (int i = 0; i < 5; i++)
            {
                int fenlei_weizhi = k;
                decimal fenlei_sum = 0;

                for (int j = 0; j < 3; j++)
                {
                    k = k + 1;
                    row1 = sheet1.CreateRow(k);                   // 原先是:row1 = sheet1.CreateRow(k+1);
                    row1.CreateCell(0).SetCellValue(laiyuanxuhao[j]);
                    row1.CreateCell(1).SetCellValue(laiyuan[j]);

                    var GenreLst = new List<shouruzongbiao>();
                    var GenreQry = from d in xiangmuzongbiaos
                                   select new shouruzongbiao
                                   {
                                       mingcheng = d.mingcheng,
                                       xiangmulaiyuan = d.xiangmulaiyuan,
                                       xiangmuleibie = d.xiangmuleibie,
                                       xiangmufenlei = d.xiangmufenlei,
                                       zijinlaiyuan = d.zijinlaiyuan,
                                       zhixingdanwei = d.zhixingdanwei,
                                       jine = d.jine,             //这里有更改
                                       xiadariqi = d.xiadariqi,
                                       zijinxingzhi = d.zijinxingzhi,
                                       xiabobumen = d.xiabobumen,

                                       niandu = d.niandu,
                                   };
                    GenreLst.AddRange(GenreQry);     //GenreQry.Distinct():去除相同的元素
                    GenreLst = GenreLst.Where(d => d.xiangmuleibie == fenlei[i] && d.zijinlaiyuan == laiyuan[j]).ToList();
                    int GenreLst_count = GenreLst.Count();

                    if (GenreLst_count == 0)
                    {
                        row1.CreateCell(9).SetCellValue(0);
                        // k = k + 1;
                    }
                    else
                    {
                        ICell cell_fenleixiaoji = row1.CreateCell(9);
                        cell_fenleixiaoji.SetCellFormula("sum(j" + (k + 2).ToString() + ":" + "j" + (k + 1 + GenreLst_count).ToString() + ")");
                        cell_fenleixiaoji.CellStyle = cellstyle_jine;
                    }

                    int h = 1;

                    foreach (var item in GenreLst)
                    //foreach (var item in zhichu)
                    {
                        row1 = sheet1.CreateRow(k + 1);

                        ICell cell = row1.CreateCell(0);
                        cell.SetCellValue(laiyuanxuhao[j] + "." + h.ToString());

                        cell = row1.CreateCell(1);
                        cell.SetCellValue(item.mingcheng);

                        cell = row1.CreateCell(2);
                        //cell.SetCellValue(item.xiangmuleibie);
                        cell.SetCellValue(item.xiangmufenlei);

                        cell = row1.CreateCell(3);
                        cell.SetCellValue(item.zijinlaiyuan);


                        cell = row1.CreateCell(4);
                        cell.SetCellValue(item.zijinxingzhi);

                        cell = row1.CreateCell(5);
                        cell.SetCellValue(item.zhixingdanwei);

                        cell = row1.CreateCell(6);
                        cell.SetCellValue(item.xiabobumen);

                        DateTime ti = DateTime.Parse(item.xiadariqi.ToString());
                        cell = row1.CreateCell(7);
                        cell.SetCellValue(ti);
                        ICellStyle cellStyle = book.CreateCellStyle();           //设置单元格日期格式，在内容上和下面的货币有些不一样，因为是2.0版本的
                        IDataFormat format = book.CreateDataFormat();
                        cellStyle.DataFormat = format.GetFormat("yyyy-m-d");
                        cell.CellStyle = cellStyle;

                        cell = row1.CreateCell(8);
                        cell.SetCellValue(item.xiangmulaiyuan);

                        fenlei_sum = fenlei_sum + item.jine;

                        float s1 = float.Parse(item.jine.ToString());
                        cell = row1.CreateCell(9);
                        cell.SetCellValue(s1);

                        //ICellStyle cellstyle = book.CreateCellStyle();                           
                        //cellstyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
                        cell.CellStyle = cellstyle_jine;

                        h = h + 1;
                        k = k + 1;
                    }
                }

                k = k + 1;

                row1 = sheet1.CreateRow(fenlei_weizhi);
                row1.CreateCell(0).SetCellValue(fenleixuhao[i]);
                row1.CreateCell(1).SetCellValue(fenlei[i]);
                row1.CreateCell(9).SetCellValue(fenlei_sum.ToString());

                zongji_sum = zongji_sum + fenlei_sum;
            }

            row_zongji.CreateCell(9).SetCellValue(zongji_sum.ToString());

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return File(ms, "application/vnd.ms-excel", "收入汇总表.xls");
        }

        [Authorize(Roles = "科长,工作人员")]
        public FileResult GetFile()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "Content/uploads/moban/";
            string fileName = "批量导入模板.xlsx";
            return File(path + fileName, "text/plain", fileName);
        }

        [Authorize(Roles = "科长,工作人员")]
        public ActionResult excelImport()
        {
            return View();
        }

        [HttpPost]
        [Authorize(Roles = "科长,工作人员")]
        public ActionResult excelImport(FormCollection form)
        {
            ViewBag.error = "";
            // ViewBag.fugai = form["gengxinmoshi"].ToString();
            HttpPostedFileBase file = Request.Files["files"];
            //string file =form["files"];

            string FileName;
            string savePath;
            if (file == null || file.ContentLength <= 0)
            {
                ViewBag.error = "文件不能为空";
                return View();
            }
            else
            {
                string filename = Path.GetFileName(file.FileName);
                int filesize = file.ContentLength;//获取上传文件的大小单位为字节byte
                string fileEx = System.IO.Path.GetExtension(filename);//获取上传文件的扩展名
                string NoFileName = System.IO.Path.GetFileNameWithoutExtension(filename);//获取无扩展名的文件名

                //int Maxsize = 4000 * 1024;//定义上传文件的最大空间大小为4M

                string FileType = ".xls,.xlsx";//定义上传文件的类型字符串

                FileName = NoFileName + DateTime.Now.ToString("yyyyMMddhhmmss") + fileEx;
                if (!FileType.Contains(fileEx))
                {
                    ViewBag.error = "文件类型不对，只能导入xls和xlsx格式的文件";
                    return View();
                }
                //if (filesize >= Maxsize)
                //{
                //    ViewBag.error = "上传文件超过4M，不能上传";
                //    return View();
                //}
                string path = AppDomain.CurrentDomain.BaseDirectory + "content/uploads/excel/";
                savePath = Path.Combine(path, FileName);
                file.SaveAs(savePath);
            }

            //string result = string.Empty;
            string strConn;
            //office 2007 可用 导入xls 不用于.xlsx
            //strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + savePath + ";" + "Extended Properties=Excel 8.0";

            strConn = "Provider=Microsoft.Ace.OleDb.12.0;Data Source=" + savePath + ";" + "Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";

            OleDbConnection conn = new OleDbConnection(strConn);
            OleDbDataAdapter myCommand = new OleDbDataAdapter("select * from [Sheet1$]", strConn);
            DataSet myDataSet = new DataSet();
            try
            {
                conn.Open();
                myCommand.Fill(myDataSet, "ExcelInfo");

            }
            catch (Exception ex)
            {
                ViewBag.error = ex.Message;
                return View(); //返回错误描述
            }
            finally
            {
                conn.Close(); //无论如何都要执行的语句。
            }

            DataTable ds = myDataSet.Tables["ExcelInfo"].DefaultView.ToTable();
            DataRow[] dr = ds.Select();


            //引用事务机制，出错时，事物回滚
            using (TransactionScope transaction = new TransactionScope())
            {
                if (form["gengxinmoshi"].ToString() == "fugai")
                {
                    string ConString = System.Configuration.ConfigurationManager.ConnectionStrings["caiwuContent"].ConnectionString;
                    SqlConnection con = new SqlConnection(ConString);
                    string sqldel = "delete from xiangmuzongbiao";
                    try
                    {
                        con.Open();

                        SqlCommand sqlcmddel = new SqlCommand(sqldel, con);

                        sqlcmddel.ExecuteNonQuery();   //删除了现有的名单

                        int j;
                        for (int i = 0; i < dr.Length; i++)
                        {
                            j = i + 1;
                            string niandu = dr[i]["年度"].ToString();                    //这里是新加上的

                            string xiangmuleibie = dr[i]["项目类别"].ToString();

                            string mingcheng = dr[i]["项目名称"].ToString();

                            string shangcaixiangmuhao = dr[i]["上财项目号"].ToString();

                            string lierulanmu = dr[i]["列入栏目"].ToString();
                            string bianhao = dr[i]["编号"].ToString();

                            string zhixingdanwei = dr[i]["执行单位"].ToString();
                            string lianxiren = dr[i]["联系人"].ToString();

                            Decimal jine;
                            if (!Decimal.TryParse(dr[i]["金额（万元）"].ToString(), out jine))
                            {
                                ViewBag.error = "第" + j + "条记录的金额格式可能有误，请认真检查后再导入！";
                                return View();
                            }

                            string xiangmufenlei = dr[i]["项目分类"].ToString();

                            string zijinlaiyuan = dr[i]["资金来源"].ToString();

                            string xiangmulaiyuan = dr[i]["项目来源"].ToString();

                            string zijinxingzhi = dr[i]["资金性质"].ToString();

                            string xiabobumen = dr[i]["下拨部门"].ToString();

                            DateTime xiadariqi;
                            if (!DateTime.TryParse(dr[i]["下达日期"].ToString(), out xiadariqi))
                            {
                                ViewBag.error = "第" + j + "条记录的下达日期格式可能有误，请认真检查后再导入！";
                                return View();
                            }

                            string xiadaqingkuang = dr[i]["下达情况"].ToString();

                            string xiaozhangbangong = dr[i]["校长办公会审议结果"].ToString();
                            string dangweihui = dr[i]["党委会审议结果"].ToString();

                            //以下是按照上面的内容进行更改的
                            //string lururen = User.Identity.Name.ToString();
                            string lururen = User.Identity.Name.ToString();

                            //string lururiqi = dr[i]["录入日期"].ToString();
                            //DateTime lururiqi = DateTime.Now;
                            //string lururiqi = dr[i]["录入日期"].ToString();

                            string lururiqi = DateTime.Now.Date.ToString("yyyy-MM-dd");

                            string shenheren = dr[i]["审核人"].ToString();

                            //DateTime shenheriqi;
                            //if (!DateTime.TryParse(dr[i]["审核日期"].ToString(), out shenheriqi))
                            //{
                            //    ViewBag.error = "第" + j + "条记录的审核日期格式可能有误，请认真检查后再导入！";
                            //    return View();
                            //}

                            string shenhezhuangtai = "未审核";

                            string beizhu = dr[i]["备注"].ToString();                    //这里是新加上的

                            string insertstr = "insert into xiangmuzongbiao (niandu, xiangmuleibie, mingcheng, shangcaixiangmuhao, lierulanmu, bianhao, zhixingdanwei,lianxiren, jine, xiangmufenlei, zijinlaiyuan, xiangmulaiyuan, zijinxingzhi,xiabobumen, xiadariqi,xiadaqingkuang, xiaozhanghui,dangweihui, lururen, lururiqi, shenheren, shenhezhuangtai ,beizhu ) values ('" + niandu + "','" + xiangmuleibie + "','" + mingcheng + "','" + shangcaixiangmuhao + "','" + lierulanmu + "','" + bianhao + "','" + zhixingdanwei + "','" + lianxiren + "','" + jine + "','" + xiangmufenlei + "','" + zijinlaiyuan + "','" + xiangmulaiyuan + "', '" + zijinxingzhi + "','" + xiabobumen + "','" + xiadariqi + "','" + xiadaqingkuang + "','" + xiaozhangbangong + "',   '" + dangweihui + "'       ,  '" + lururen + "','" + lururiqi + "','" + shenheren + "', '" + shenhezhuangtai + "','" + beizhu + "')";//'" + shenheriqi + "',

                            SqlCommand cmd = new SqlCommand(insertstr, con);

                            try
                            {
                                cmd.ExecuteNonQuery();
                            }

                            catch (Exception ex)
                            {
                                ViewBag.error = "第" + j + "条记录插入有误，请认真检查格式后再导入！覆盖";

                                return View();
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.error = ex.Message;
                        return View(); //返回错误描述
                    }
                    finally
                    {
                        con.Close(); //无论如何都要执行的语句。
                    }
                    // ViewBag.fugai = "离开覆盖了";
                }
                else  //追加模式
                {

                    string ConString = System.Configuration.ConfigurationManager.ConnectionStrings["caiwuContent"].ConnectionString;
                    SqlConnection con = new SqlConnection(ConString);

                    try
                    {
                        con.Open();
                        int j;
                        for (int i = 0; i < dr.Length; i++)
                        {
                            j = i + 1;

                            string niandu = dr[i]["年度"].ToString();                    //这里是新加上的

                            string xiangmuleibie = dr[i]["项目类别"].ToString();

                            string mingcheng = dr[i]["项目名称"].ToString();

                            string shangcaixiangmuhao = dr[i]["上财项目号"].ToString();

                            string lierulanmu = dr[i]["列入栏目"].ToString();
                            string bianhao = dr[i]["编号"].ToString();

                            string zhixingdanwei = dr[i]["执行单位"].ToString();
                            string lianxiren = dr[i]["联系人"].ToString();

                            Decimal jine;
                            if (!Decimal.TryParse(dr[i]["金额（万元）"].ToString(), out jine))
                            {
                                ViewBag.error = "第" + j + "条记录的金额格式可能有误，请认真检查后再导入！";
                                return View();
                            }

                            string xiangmufenlei = dr[i]["项目分类"].ToString();

                            string zijinlaiyuan = dr[i]["资金来源"].ToString();

                            string xiangmulaiyuan = dr[i]["项目来源"].ToString();

                            string zijinxingzhi = dr[i]["资金性质"].ToString();

                            string xiabobumen = dr[i]["下拨部门"].ToString();

                            DateTime xiadariqi;
                            if (!DateTime.TryParse(dr[i]["下达日期"].ToString(), out xiadariqi))
                            {
                                ViewBag.error = "第" + j + "条记录的下达日期格式可能有误，请认真检查后再导入！";
                                return View();
                            }

                            string xiadaqingkuang = dr[i]["下达情况"].ToString();

                            string xiaozhangbangong = dr[i]["校长办公会审议结果"].ToString();
                            string dangweihui = dr[i]["党委会审议结果"].ToString();

                            //以下是按照上面的内容进行更改的
                            //string lururen = User.Identity.Name.ToString();
                            string lururen = User.Identity.Name.ToString();

                            //string lururiqi = dr[i]["录入日期"].ToString();
                            //DateTime lururiqi = DateTime.Now;
                            //string lururiqi = dr[i]["录入日期"].ToString();

                            string lururiqi = DateTime.Now.Date.ToString("yyyy-MM-dd");

                            string shenheren = dr[i]["审核人"].ToString();

                            //DateTime shenheriqi;
                            //if (!DateTime.TryParse(dr[i]["审核日期"].ToString(), out shenheriqi))
                            //{
                            //    ViewBag.error = "第" + j + "条记录的审核日期格式可能有误，请认真检查后再导入！";
                            //    return View();
                            //}

                            string shenhezhuangtai = "未审核";

                            string beizhu = dr[i]["备注"].ToString();                    //这里是新加上的

                            string insertstr = "insert into xiangmuzongbiao (niandu, xiangmuleibie, mingcheng, shangcaixiangmuhao, lierulanmu, bianhao, zhixingdanwei,lianxiren, jine, xiangmufenlei, zijinlaiyuan, xiangmulaiyuan, zijinxingzhi,xiabobumen, xiadariqi,xiadaqingkuang, xiaozhanghui,dangweihui, lururen, lururiqi, shenheren, shenhezhuangtai ,beizhu ) values ('" + niandu + "','" + xiangmuleibie + "','" + mingcheng + "','" + shangcaixiangmuhao + "','" + lierulanmu + "','" + bianhao + "','" + zhixingdanwei + "','" + lianxiren + "','" + jine + "','" + xiangmufenlei + "','" + zijinlaiyuan + "','" + xiangmulaiyuan + "', '" + zijinxingzhi + "','" + xiabobumen + "','" + xiadariqi + "','" + xiadaqingkuang + "','" + xiaozhangbangong + "',   '" + dangweihui + "'       ,  '" + lururen + "','" + lururiqi + "','" + shenheren + "', '" + shenhezhuangtai + "','" + beizhu + "')";//'" + shenheriqi + "',


                            SqlCommand cmd = new SqlCommand(insertstr, con);

                            try
                            {
                                cmd.ExecuteNonQuery();
                            }

                            catch (Exception ex)
                            {
                                ViewBag.error = "第" + j + "条记录插入有误，请认真检查格式后再导入！追加";

                                return View();
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        ViewBag.error = ex.Message;
                        return View(); //返回错误描述
                    }
                    finally
                    {
                        con.Close(); //无论如何都要执行的语句。
                    }

                }
                transaction.Complete();
            }
            ViewBag.error = "祝贺您，本次企业信息导入成功！";
            System.Threading.Thread.Sleep(2000);
            return RedirectToAction("./index");
        }

        //[HttpGet]
        public async Task<ActionResult> doczhichuv(string xmleibie, int? niandu)
        {
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  select m;

            if (niandu.HasValue == true)
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == niandu && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == jinnian && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = jinnian;
            }

            if (String.IsNullOrEmpty(xmleibie) || xmleibie == "全部")
            {
                ViewBag.leibie = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.xiangmuleibie == xmleibie);
                ViewBag.leibie = xmleibie;
            }
            var GenreLst = new List<zhichuzongbiao>();
            var GenreQry = from d in xiangmuzongbiao
                           select new zhichuzongbiao
                           {
                               ID = d.ID,
                               shangcaixiangmuhao = d.shangcaixiangmuhao,
                               mingcheng = d.mingcheng,

                               lierulanmu = d.lierulanmu,
                               xiangmuleibie = d.xiangmuleibie,
                               xiangmulaiyuan = d.xiangmulaiyuan,
                               zhixingdanwei = d.zhixingdanwei,
                               lianxiren = d.lianxiren,
                               jine = d.jine,

                               xiangmufenlei = d.xiangmufenlei,                //这是新加上的
                               zijinlaiyuan = d.zijinlaiyuan,

                               shenheren = d.shenheren,

                               shenhezhuangtai = d.shenhezhuangtai,

                               xiabobumen = d.xiabobumen,
                               niandu = d.niandu,
                           };
            GenreLst.AddRange(GenreQry.Distinct());     //GenreQry.Distinct():去除相同的元素
            int l = GenreLst.Count();
            return View(await GenreQry.ToListAsync());
        }

        //如何字符串长度超过juli个英文字母就插入一个换行符
        public string inserthuiche(string str, int juli)
        {
            if (!String.IsNullOrEmpty(str))
            {
                int len_temp = 0;
                int n = str.Length;

                string strWord = "";
                int asc;
                for (int i = 0; i < n; i++)
                {
                    strWord = str.Substring(i, 1);
                    asc = Convert.ToChar(strWord);
                    if (asc < 0 || asc > 127)
                        len_temp = len_temp + 2;
                    else
                        len_temp = len_temp + 1;

                    if (len_temp >= juli)
                    {
                        len_temp = 0;
                        str = str.Insert(i + 1, "\n");
                    }
                }
            }
            return str;
        }
        

        //20150317这里更改，导出doc
        //[HttpPost]
        public FileResult doczhichu(string xmleibie, int? niandu)
        {           
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  select m;

            if (niandu.HasValue == true)
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == niandu && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == jinnian && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = jinnian;
            }
            
            if (String.IsNullOrEmpty(xmleibie) || xmleibie == "全部")
            {
                ViewBag.leibie = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.xiangmuleibie == xmleibie);
                ViewBag.leibie = xmleibie;
            }

            var GenreLst = new List<zhichuzongbiao>();
            var GenreQry = from d in xiangmuzongbiao
                           select new zhichuzongbiao
                           {
                               shangcaixiangmuhao = d.shangcaixiangmuhao,
                               mingcheng = d.mingcheng,

                               lierulanmu = d.lierulanmu,
                               xiangmulaiyuan = d.xiangmulaiyuan,
                               zhixingdanwei = d.zhixingdanwei,
                               lianxiren = d.lianxiren,
                               jine = d.jine,

                               xiangmufenlei = d.xiangmuleibie,                //这是新加上的
                               zijinlaiyuan = d.zijinlaiyuan,

                               shenheren = d.shenheren,

                               shenhezhuangtai = d.shenhezhuangtai,

                               xiabobumen = d.xiabobumen,
                               niandu = d.niandu,
                               xiaozhanghui = d.xiaozhanghui,
                               dangweihui = d.dangweihui,
                               shenpiliucheng = d.shenpiliucheng,
                           };
            GenreLst.AddRange(GenreQry.Distinct());     //GenreQry.Distinct():去除相同的元素            

            string filename = Server.MapPath("~/Content/kong.docx");

            string outfilepath = Server.MapPath("~/Content/kong1.docx");

            string niandu_biaoming = niandu.ToString();

            DocX doc;
            //doc = DocX.Create(filename);
            doc = DocX.Load(filename);
            try
            {
                int k = 0;
                int shuzi = GenreLst.Count();

                foreach (var item in GenreLst)
                {

                    Paragraph p3 = doc.InsertParagraph();
                    p3.Append(item.niandu + "年追加预算情况表");
                    p3.FontSize(28);
                    p3.Alignment = Alignment.center;

                    Paragraph p4 = doc.InsertParagraph();
                    p4.Append("***财预追" + item.niandu + "（        ）号");
                    p4.FontSize(12);
                    p4.Alignment = Alignment.right;

                    Table table1 = doc.AddTable(9, 4);

                    p4.InsertTableAfterSelf(table1);            //在它之后插入一个表格
                    table1.Alignment = Alignment.center;
                    table1.Paragraphs[0].FontSize(12);
                    table1.Design = TableDesign.TableGrid;

                    table1.Rows[0].Cells[0].Paragraphs[0].Append("项目名称");
                    table1.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[2].Paragraphs[0].Append("项目金额\n（万元）");
                    table1.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[0].Width = 103;

                    table1.Rows[0].Cells[1].Width = 361;
                    table1.Rows[0].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[1].Paragraphs[0].Append(inserthuiche(item.mingcheng, 26));
                    table1.Rows[0].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[3].Paragraphs[0].Append(item.jine.ToString());
                    table1.Rows[0].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[1].Paragraphs[0].FontSize(12);
                    //table1.Rows[0].Cells[0].MarginTop = 7;

                    table1.Rows[1].Cells[0].Paragraphs[0].Append("执行单位\n联系人");
                    table1.Rows[1].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[2].Paragraphs[0].Append("追加金额\n（万元）");
                    table1.Rows[1].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[1].Paragraphs[0].Append(item.zhixingdanwei + item.lianxiren);
                    table1.Rows[1].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[3].Paragraphs[0].Append("");
                    table1.Rows[1].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[3].Paragraphs[0].FontSize(12);
                    //table1.Rows[1].Cells[0].MarginTop = 7;

                    table1.Rows[2].Cells[0].Paragraphs[0].Append("列入栏目");
                    table1.Rows[2].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[2].Paragraphs[0].Append("资金来源");
                    table1.Rows[2].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[1].Paragraphs[0].Append(inserthuiche(item.lierulanmu, 26));
                    table1.Rows[2].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[3].Paragraphs[0].Append(item.zijinlaiyuan);
                    table1.Rows[2].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[3].Paragraphs[0].FontSize(12);
                    //table1.Rows[2].Cells[0].MarginTop = 7;
                    table1.Rows[2].Height = 40;

                    table1.Rows[3].Cells[0].Paragraphs[0].Append("审批流程");
                    table1.Rows[3].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[3].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].MergeCells(1, 3);
                    //table1.Rows[3].Cells[0].Width = 80;
                    //table1.Rows[3].Cells[1].Width = 830;
                    table1.Rows[3].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Height = 40;
                    //table1.Rows[3].Cells[1].Paragraphs[0].Append(item.shenheren + item.shenhezhuangtai);
                    table1.Rows[3].Cells[1].Paragraphs[0].Append(item.shenpiliucheng);
                    table1.Rows[3].Cells[1].Paragraphs[0].FontSize(12);
                    //table1.Rows[3].Cells[0].MarginTop = 10;

                    table1.Rows[4].Cells[0].Paragraphs[0].Append("下达部门");
                    table1.Rows[4].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[4].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].MergeCells(1, 3);
                    table1.Rows[4].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[4].Cells[0].Width = 25;
                    table1.Rows[4].Height = 40;
                    table1.Rows[4].Cells[1].Paragraphs[0].Append(item.xiabobumen);
                    table1.Rows[4].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[1].Paragraphs[0].FontSize(12);
                    //table1.Rows[4].Cells[0].MarginTop = 10;

                    table1.Rows[5].Height = 130;
                    table1.Rows[5].Cells[0].Paragraphs[0].Append("校长办公会\n审议结果");
                    table1.Rows[5].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[5].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[5].MergeCells(1, 3);
                    table1.Rows[5].Cells[0].Paragraphs[0].FontSize(12);
                    if (item.xiaozhanghui == "同意立项")
                    {
                        table1.Rows[5].Cells[1].Paragraphs[0].Append("☑1.同意立项             □2.不同意立项");
                    }
                    else
                    {
                        table1.Rows[5].Cells[1].Paragraphs[0].Append("□1.同意立项             ☑2.不同意立项");
                    }
                    table1.Rows[5].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[5].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;

                    table1.Rows[5].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[6].Cells[0].Paragraphs[0].Append("党委会\n审议结果");
                    table1.Rows[6].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[6].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[6].MergeCells(1, 3);
                    table1.Rows[6].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[6].Cells[0].Width = 25;
                    table1.Rows[6].Height = 130;
                    if (item.dangweihui == "同意立项")
                    {
                        table1.Rows[6].Cells[1].Paragraphs[0].Append("☑1.同意立项             □2.不同意立项");
                    }
                    else
                    {
                        table1.Rows[6].Cells[1].Paragraphs[0].Append("□1.同意立项             ☑2.不同意立项");
                    }
                    table1.Rows[6].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[6].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[6].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;
                    //table1.Rows[6].Cells[0].MarginTop = 50;

                    table1.Rows[7].Cells[0].Paragraphs[0].Append("预算科审核");
                    table1.Rows[7].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[7].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[7].MergeCells(1, 3);
                    table1.Rows[7].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[7].Cells[0].Width = 25;
                    table1.Rows[7].Height = 120;
                    table1.Rows[7].Cells[1].Paragraphs[0].Append("签名(盖章)                年    月    日");
                    table1.Rows[7].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[7].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[7].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;
                    //table1.Rows[7].Cells[0].MarginTop = 30;
                    //table1.Rows[7].Cells[1].Paragraphs[1].AppendLine();


                    table1.Rows[8].Cells[0].Paragraphs[0].Append("处长审核");
                    table1.Rows[8].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[8].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[8].MergeCells(1, 3);
                    table1.Rows[8].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[8].Cells[0].Width = 25;
                    table1.Rows[8].Height = 120;
                    table1.Rows[8].Cells[1].Paragraphs[0].Append("签名(盖章)                年    月    日");

                    table1.Rows[8].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[8].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[8].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;
                    //table1.Rows[8].Cells[0].MarginTop = 30;
                    //table1.Rows[8].Cells[1].Paragraphs[1].AppendLine();

                    k++;
                    if (k < GenreLst.Count())
                    {
                        table1.InsertPageBreakAfterSelf();
                    }
                }
                //  doc.SaveAs(@"C:\Users\lu\Desktop\haishicaiwudoc20150203\haishicaiwuchu20150105\yusuanxiangmu\Content\outfiles\kong1.docx");
                doc.SaveAs(outfilepath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return File(outfilepath, "application/ms-word", "zhuijiamingxibiao.docx");
        }

        public FileResult doczhichu_danxiang(int id)
        {
            var xiangmuzongbiaos = from m in db.xiangmuzongbiaos
                                   where m.ID == id
                                   select m;
            var GenreLst = new List<zhichuzongbiao>();
            var GenreQry = from d in xiangmuzongbiaos
                           select new zhichuzongbiao
                           {
                               shangcaixiangmuhao = d.shangcaixiangmuhao,
                               mingcheng = d.mingcheng,

                               lierulanmu = d.lierulanmu,
                               xiangmulaiyuan = d.xiangmulaiyuan,
                               zhixingdanwei = d.zhixingdanwei,
                               lianxiren = d.lianxiren,
                               jine = d.jine,

                               xiangmufenlei = d.xiangmuleibie,                //这是新加上的
                               zijinlaiyuan = d.zijinlaiyuan,

                               shenheren = d.shenheren,

                               shenhezhuangtai = d.shenhezhuangtai,

                               xiabobumen = d.xiabobumen,
                               niandu = d.niandu,
                               xiaozhanghui = d.xiaozhanghui,
                               dangweihui = d.dangweihui,
                               shenpiliucheng = d.shenpiliucheng,
                           };
            GenreLst.AddRange(GenreQry.Distinct());     //GenreQry.Distinct():去除相同的元素
            
            string filename = Server.MapPath("~/Content/kong.docx");

            string outfilepath = Server.MapPath("~/Content/kong1.docx");

            DocX doc;
            //doc = DocX.Create(filename);
            doc = DocX.Load(filename);
            try
            {
                int k = 0;

                int shuzi = GenreLst.Count();

                foreach (var item in GenreLst)
                {
                    Paragraph p3 = doc.InsertParagraph();
                    p3.Append(item.niandu + "年追加预算情况表");
                    p3.FontSize(28);
                    p3.Alignment = Alignment.center;

                    Paragraph p4 = doc.InsertParagraph();
                    p4.Append("海师财预追" + item.niandu + "（        ）号");
                    p4.FontSize(12);
                    p4.Alignment = Alignment.right;

                    Table table1 = doc.AddTable(9, 4);

                    p4.InsertTableAfterSelf(table1);            //在它之后插入一个表格
                    table1.Alignment = Alignment.center;
                    table1.Paragraphs[0].FontSize(12);
                    table1.Design = TableDesign.TableGrid;

                    table1.Rows[0].Cells[0].Paragraphs[0].Append("项目名称");
                    table1.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[2].Paragraphs[0].Append("项目金额\n（万元）");
                    table1.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[0].Width = 103;

                    table1.Rows[0].Cells[1].Width = 361;
                    table1.Rows[0].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[1].Paragraphs[0].Append(inserthuiche(item.mingcheng, 26));
                    table1.Rows[0].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[3].Paragraphs[0].Append(item.jine.ToString());
                    table1.Rows[0].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[1].Paragraphs[0].FontSize(12);
                    //table1.Rows[0].Cells[0].MarginTop = 7;

                    table1.Rows[1].Cells[0].Paragraphs[0].Append("执行单位\n联系人");
                    table1.Rows[1].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[2].Paragraphs[0].Append("追加金额\n（万元）");
                    table1.Rows[1].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[1].Paragraphs[0].Append(item.zhixingdanwei + item.lianxiren);
                    table1.Rows[1].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[3].Paragraphs[0].Append("");
                    table1.Rows[1].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[3].Paragraphs[0].FontSize(12);
                    //table1.Rows[1].Cells[0].MarginTop = 7;

                    table1.Rows[2].Cells[0].Paragraphs[0].Append("列入栏目");
                    table1.Rows[2].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[2].Paragraphs[0].Append("资金来源");
                    table1.Rows[2].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[1].Paragraphs[0].Append(inserthuiche(item.lierulanmu, 26));
                    table1.Rows[2].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[3].Paragraphs[0].Append(item.zijinlaiyuan);
                    table1.Rows[2].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[3].Paragraphs[0].FontSize(12);
                    //table1.Rows[2].Cells[0].MarginTop = 7;
                    table1.Rows[2].Height = 40;

                    table1.Rows[3].Cells[0].Paragraphs[0].Append("审批流程");
                    table1.Rows[3].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[3].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].MergeCells(1, 3);
                    //table1.Rows[3].Cells[0].Width = 80;
                    //table1.Rows[3].Cells[1].Width = 830;
                    table1.Rows[3].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Height = 40;
                    //table1.Rows[3].Cells[1].Paragraphs[0].Append(item.shenheren + item.shenhezhuangtai);
                    table1.Rows[3].Cells[1].Paragraphs[0].Append(item.shenpiliucheng);
                    table1.Rows[3].Cells[1].Paragraphs[0].FontSize(12);
                    //table1.Rows[3].Cells[0].MarginTop = 10;

                    table1.Rows[4].Cells[0].Paragraphs[0].Append("下达部门");
                    table1.Rows[4].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[4].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].MergeCells(1, 3);
                    table1.Rows[4].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[4].Cells[0].Width = 25;
                    table1.Rows[4].Height = 40;
                    table1.Rows[4].Cells[1].Paragraphs[0].Append(item.xiabobumen);
                    table1.Rows[4].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[1].Paragraphs[0].FontSize(12);
                    //table1.Rows[4].Cells[0].MarginTop = 10;

                    table1.Rows[5].Height = 130;
                    table1.Rows[5].Cells[0].Paragraphs[0].Append("校长办公会\n审议结果");
                    table1.Rows[5].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[5].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[5].MergeCells(1, 3);
                    table1.Rows[5].Cells[0].Paragraphs[0].FontSize(12);
                    if (item.xiaozhanghui == "同意立项")
                    {
                        table1.Rows[5].Cells[1].Paragraphs[0].Append("☑1.同意立项             □2.不同意立项");
                    }
                    else
                    {
                        table1.Rows[5].Cells[1].Paragraphs[0].Append("□1.同意立项             ☑2.不同意立项");
                    }
                    table1.Rows[5].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[5].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;

                    table1.Rows[5].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[6].Cells[0].Paragraphs[0].Append("党委会\n审议结果");
                    table1.Rows[6].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[6].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[6].MergeCells(1, 3);
                    table1.Rows[6].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[6].Cells[0].Width = 25;
                    table1.Rows[6].Height = 130;
                    if (item.dangweihui == "同意立项")
                    {
                        table1.Rows[6].Cells[1].Paragraphs[0].Append("☑1.同意立项             □2.不同意立项");
                    }
                    else
                    {
                        table1.Rows[6].Cells[1].Paragraphs[0].Append("□1.同意立项             ☑2.不同意立项");
                    }
                    table1.Rows[6].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[6].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[6].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;
                    //table1.Rows[6].Cells[0].MarginTop = 50;

                    table1.Rows[7].Cells[0].Paragraphs[0].Append("预算科审核");
                    table1.Rows[7].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[7].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[7].MergeCells(1, 3);
                    table1.Rows[7].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[7].Cells[0].Width = 25;
                    table1.Rows[7].Height = 120;
                    table1.Rows[7].Cells[1].Paragraphs[0].Append("签名(盖章)                年    月    日");
                    table1.Rows[7].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[7].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[7].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;
                    //table1.Rows[7].Cells[0].MarginTop = 30;
                    //table1.Rows[7].Cells[1].Paragraphs[1].AppendLine();

                    table1.Rows[8].Cells[0].Paragraphs[0].Append("处长审核");
                    table1.Rows[8].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[8].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[8].MergeCells(1, 3);
                    table1.Rows[8].Cells[0].Paragraphs[0].FontSize(12);
                    //table1.Rows[8].Cells[0].Width = 25;
                    table1.Rows[8].Height = 120;
                    table1.Rows[8].Cells[1].Paragraphs[0].Append("签名(盖章)                年    月    日");

                    table1.Rows[8].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[8].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[8].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Bottom;
                    //table1.Rows[8].Cells[0].MarginTop = 30;
                    //table1.Rows[8].Cells[1].Paragraphs[1].AppendLine();

                    k++;
                    if (k < GenreLst.Count())
                    {
                        table1.InsertPageBreakAfterSelf();
                    }
                }
                //  doc.SaveAs(@"C:\Users\lu\Desktop\haishicaiwudoc20150203\haishicaiwuchu20150105\yusuanxiangmu\Content\outfiles\kong1.docx");
                doc.SaveAs(outfilepath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return File(outfilepath, "application/ms-word", "zhuijiamingxibiao.docx");
        }

        public async Task<ActionResult> docshouruv(string xmleibie, int? niandu)
        {
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  select m;

            if (niandu.HasValue == true)
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == niandu && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = niandu;
            }
            else
            {
                int jinnian = int.Parse(DateTime.Now.ToString("yyyy"));
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.niandu == jinnian && s.shenhezhuangtai == "通过");
                ViewBag.nianduzhi = jinnian;
            }

            if (String.IsNullOrEmpty(xmleibie) || xmleibie == "全部")
            {
                ViewBag.leibie = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.xiangmuleibie == xmleibie);
                ViewBag.leibie = xmleibie;
            }
            var GenreLst = new List<shouruzongbiao>();
            var GenreQry = from d in xiangmuzongbiao
                           select new shouruzongbiao
                           {
                               ID = d.ID,
                               mingcheng = d.mingcheng,
                               xiangmuleibie = d.xiangmuleibie,
                               xiangmulaiyuan = d.xiangmulaiyuan,
                               xiangmufenlei = d.xiangmufenlei,
                               shancaixiangmuhao = d.shangcaixiangmuhao,


                               zijinlaiyuan = d.zijinlaiyuan,
                               zhixingdanwei = d.zhixingdanwei,
                               jine = d.jine,             //这里有更改
                               xiadariqi = d.xiadariqi,
                               zijinxingzhi = d.zijinxingzhi,
                               xiabobumen = d.xiabobumen,
                               niandu = d.niandu,
                           };
            GenreLst.AddRange(GenreQry.Distinct());     //GenreQry.Distinct():去除相同的元素

            return View(await GenreQry.ToListAsync());
        }

        public FileResult docshouru(string xmleibie, int? niandu)
        {
            var xiangmuzongbiao = from m in db.xiangmuzongbiaos
                                  where m.shenhezhuangtai == "通过" && m.niandu == niandu
                                  select m;
            if (String.IsNullOrEmpty(xmleibie) || xmleibie == "全部")
            {
                ViewBag.leibie = "全部";
            }
            else
            {
                xiangmuzongbiao = xiangmuzongbiao.Where(s => s.xiangmuleibie == xmleibie);
                ViewBag.leibie = xmleibie;
            }
            var GenreLst = new List<shouruzongbiao>();
            var GenreQry = from d in xiangmuzongbiao
                           select new shouruzongbiao
                           {
                               mingcheng = d.mingcheng,
                               xiangmulaiyuan = d.xiangmulaiyuan,
                               xiangmufenlei = d.xiangmufenlei,
                               zijinlaiyuan = d.zijinlaiyuan,
                               zhixingdanwei = d.zhixingdanwei,
                               shancaixiangmuhao = d.shangcaixiangmuhao,
                               jine = d.jine,             //这里有更改
                               xiadariqi = d.xiadariqi,
                               zijinxingzhi = d.zijinxingzhi,
                               xiabobumen = d.xiabobumen,
                               niandu = d.niandu,
                           };
            GenreLst.AddRange(GenreQry.Distinct());     //GenreQry.Distinct():去除相同的元素

            string filename = Server.MapPath("~/Content/shouru.docx");

            string outfilepath = Server.MapPath("~/Content/shouru1.docx");

            DocX doc;
            doc = DocX.Load(filename);

            string niandu_biaoming = niandu.ToString();

            try
            {
                int k = 0;

                foreach (var item in GenreLst)
                {
                    Paragraph p3 = doc.InsertParagraph();
                    p3.Append(item.niandu + "年收入预算明细表");
                    p3.FontSize(28);
                    p3.Alignment = Alignment.center;

                    //p4.AppendLine();
                    Paragraph p4 = doc.InsertParagraph();
                    p4.Append("海师财预收" + item.niandu + "（        ）号");
                    p4.FontSize(12);
                    p4.Alignment = Alignment.right;

                    Table table1 = doc.AddTable(5, 4);
                    p4.InsertTableAfterSelf(table1);            //在它之后插入一个表格
                                                                // table1.Alignment = Alignment.center;
                    table1.Paragraphs[0].FontSize(12);
                    table1.Design = TableDesign.TableGrid;

                    table1.Rows[0].Cells[0].Width = 15;
                    table1.Rows[0].Cells[1].Width = 450;
                    table1.Rows[0].Cells[2].Width = 15;
                    table1.Rows[0].Cells[3].Width = 150;

                    table1.Rows[0].Cells[0].Paragraphs[0].Append("项目名称");
                    table1.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[2].Paragraphs[0].Append("项目金额\n（万元）");
                    table1.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[1].Paragraphs[0].Append(item.mingcheng);
                    table1.Rows[0].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[3].Paragraphs[0].Append(item.jine.ToString());
                    table1.Rows[0].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[3].Paragraphs[0].FontSize(12);

                    //table1.Rows[0].Cells[3].Width = 80;
                    //table1.Rows[0].Cells[0].MarginTop = 10;
                    table1.Rows[0].Height = 80;

                    table1.Rows[1].Cells[0].Paragraphs[0].Append("项目号"); //在数据库中对应的是“上财项目号”
                    table1.Rows[1].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[2].Paragraphs[0].Append("项目分类");
                    table1.Rows[1].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[2].Paragraphs[0].FontSize(12);

                    table1.Rows[1].Cells[1].Paragraphs[0].Append(item.shancaixiangmuhao);
                    table1.Rows[1].Cells[1].MarginLeft = 0;
                    table1.Rows[1].Cells[1].MarginRight = 0;
                    table1.Rows[1].Cells[0].Paragraphs[0].Alignment = Alignment.left;
                    table1.Rows[1].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;

                    table1.Rows[1].Cells[3].Paragraphs[0].Append(item.xiangmufenlei);
                    table1.Rows[1].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Height = 80;
                    //table1.Rows[1].Cells[0].MarginTop = 10;

                    table1.Rows[2].Cells[0].Paragraphs[0].Append("项目来源");
                    table1.Rows[2].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[2].Paragraphs[0].Append("资金来源");
                    table1.Rows[2].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[1].Paragraphs[0].Append(item.xiangmulaiyuan);
                    table1.Rows[2].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[3].Paragraphs[0].Append(item.zijinlaiyuan);
                    table1.Rows[2].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Height = 80;
                    //table1.Rows[2].Cells[0].MarginTop = 15;

                    table1.Rows[3].Cells[0].Paragraphs[0].Append("资金性质");
                    table1.Rows[3].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[3].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[2].Paragraphs[0].Append("执行单位");
                    table1.Rows[3].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[3].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Cells[1].Paragraphs[0].Append(item.zijinxingzhi);
                    table1.Rows[3].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[3].Paragraphs[0].Append(item.zhixingdanwei);
                    table1.Rows[3].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Height = 80;
                    //table1.Rows[3].Cells[0].MarginTop = 15;

                    table1.Rows[4].Cells[0].Paragraphs[0].Append("下拨部门");
                    table1.Rows[4].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[4].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[2].Paragraphs[0].Append("下达日期");
                    table1.Rows[4].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[4].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Cells[1].Paragraphs[0].Append(item.zijinxingzhi);
                    table1.Rows[4].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[3].Paragraphs[0].Append(item.xiadariqi.ToShortDateString());
                    table1.Rows[4].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Height = 80;
                    
                    k++;
                    if (k < GenreLst.Count())
                    {
                        table1.InsertPageBreakAfterSelf();
                    }
                }
                doc.SaveAs(outfilepath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return File(outfilepath, "application/ms-word", "shourumingxibiao.docx");
        }


        public FileResult docshouru_danxiang(int id)
        {
            var xiangmuzongbiaos = from m in db.xiangmuzongbiaos
                                   where m.ID == id
                                   select m;
            var GenreLst = new List<shouruzongbiao>();
            var GenreQry = from d in xiangmuzongbiaos
                           select new shouruzongbiao
                           {
                               mingcheng = d.mingcheng,
                               xiangmulaiyuan = d.xiangmulaiyuan,
                               xiangmufenlei = d.xiangmufenlei,
                               zijinlaiyuan = d.zijinlaiyuan,
                               zhixingdanwei = d.zhixingdanwei,
                               jine = d.jine,             //这里有更改
                               shancaixiangmuhao = d.shangcaixiangmuhao,
                               xiadariqi = d.xiadariqi,
                               zijinxingzhi = d.zijinxingzhi,
                               xiabobumen = d.xiabobumen,
                               niandu = d.niandu,

                           };
            GenreLst.AddRange(GenreQry.Distinct());     //GenreQry.Distinct():去除相同的元素

            int zongshu = GenreLst.Count();
            string filename = Server.MapPath("~/Content/shouru.docx");
            string outfilepath = Server.MapPath("~/Content/shouru1.docx");

            DocX doc;
            doc = DocX.Load(filename);
            //string niandu_biaoming = item.niandu.ToString();
            try
            {
                int k = 0;
                foreach (var item in GenreLst)
                {
                    Paragraph p3 = doc.InsertParagraph();
                    p3.Append(item.niandu + "年收入预算明细表");
                    p3.FontSize(28);
                    p3.Alignment = Alignment.center;

                    //p4.AppendLine();
                    Paragraph p4 = doc.InsertParagraph();
                    p4.Append("***财预收" + item.niandu + "（        ）号");
                    p4.FontSize(12);
                    p4.Alignment = Alignment.right;

                    Table table1 = doc.AddTable(5, 4);
                    p4.InsertTableAfterSelf(table1);            //在它之后插入一个表格
                                                                // table1.Alignment = Alignment.center;
                    table1.Paragraphs[0].FontSize(12);
                    table1.Design = TableDesign.TableGrid;

                    table1.Rows[0].Cells[0].Width = 15;
                    table1.Rows[0].Cells[1].Width = 450;
                    table1.Rows[0].Cells[2].Width = 15;
                    table1.Rows[0].Cells[3].Width = 150;


                    table1.Rows[0].Cells[0].Paragraphs[0].Append("项目名称");
                    table1.Rows[0].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[2].Paragraphs[0].Append("项目金额\n（万元）");
                    table1.Rows[0].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[0].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[1].Paragraphs[0].Append(item.mingcheng);
                    table1.Rows[0].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[3].Paragraphs[0].Append(item.jine.ToString());
                    table1.Rows[0].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[0].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[0].Cells[3].Paragraphs[0].FontSize(12);

                    //table1.Rows[0].Cells[3].Width = 80;
                    //table1.Rows[0].Cells[0].MarginTop = 10;
                    table1.Rows[0].Height = 80;

                    table1.Rows[1].Cells[0].Paragraphs[0].Append("项目号"); //在数据库中对应的是“上财项目号”
                    table1.Rows[1].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[2].Paragraphs[0].Append("项目分类");
                    table1.Rows[1].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[1].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[2].Paragraphs[0].FontSize(12);

                    table1.Rows[1].Cells[1].Paragraphs[0].Append(item.shancaixiangmuhao);
                    table1.Rows[1].Cells[1].MarginLeft = 0;
                    table1.Rows[1].Cells[1].MarginRight = 0;
                    table1.Rows[1].Cells[0].Paragraphs[0].Alignment = Alignment.left;
                    table1.Rows[1].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;

                    table1.Rows[1].Cells[3].Paragraphs[0].Append(item.xiangmufenlei);
                    table1.Rows[1].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[1].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[1].Height = 80;
                    //table1.Rows[1].Cells[0].MarginTop = 10;

                    table1.Rows[2].Cells[0].Paragraphs[0].Append("项目来源");
                    table1.Rows[2].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[2].Paragraphs[0].Append("资金来源");
                    table1.Rows[2].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[2].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[1].Paragraphs[0].Append(item.xiangmulaiyuan);
                    table1.Rows[2].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[3].Paragraphs[0].Append(item.zijinlaiyuan);
                    table1.Rows[2].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[2].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[2].Height = 80;
                    //table1.Rows[2].Cells[0].MarginTop = 15;

                    table1.Rows[3].Cells[0].Paragraphs[0].Append("资金性质");
                    table1.Rows[3].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[3].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[2].Paragraphs[0].Append("执行单位");
                    table1.Rows[3].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[3].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Cells[1].Paragraphs[0].Append(item.zijinxingzhi);
                    table1.Rows[3].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[3].Paragraphs[0].Append(item.zhixingdanwei);
                    table1.Rows[3].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[3].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[3].Height = 80;
                    //table1.Rows[3].Cells[0].MarginTop = 15;

                    table1.Rows[4].Cells[0].Paragraphs[0].Append("下拨部门");
                    table1.Rows[4].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[4].Cells[0].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[2].Paragraphs[0].Append("下达日期");
                    table1.Rows[4].Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    table1.Rows[4].Cells[2].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[0].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Cells[2].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Cells[1].Paragraphs[0].Append(item.zijinxingzhi);
                    table1.Rows[4].Cells[1].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[3].Paragraphs[0].Append(item.xiadariqi.ToShortDateString());
                    table1.Rows[4].Cells[3].VerticalAlignment = Novacode.VerticalAlignment.Center;
                    table1.Rows[4].Cells[1].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Cells[3].Paragraphs[0].FontSize(12);
                    table1.Rows[4].Height = 80;
                    //table1.Rows[4].Cells[0].MarginTop = 15;

                    //Table table1 = doc.AddTable(2, 2);
                    //table1.Design = TableDesign.ColorfulGridAccent2;
                    //table1.Alignment = Alignment.center;

                    //table1.Rows[1].Cells[0].Width = 200;
                    //table1.Rows[1].Height = 200;
                    //table1.Rows[1].Cells[1].Width = 800;
                    //table1.Rows[0].Cells[0].Paragraphs[0].Append("党委会党委会党委会的方位的房间打开");
                    //table1.Rows[0].Cells[1].Paragraphs[0].Append("2");
                    //table1.Rows[1].MergeCells(0, 1);

                    //table1.Rows[1].Cells[0].Paragraphs[0].Append("4");

                    //p4.InsertTableBeforeSelf(table1);

                    //p4.InsertPageBreakAfterSelf();

                    k++;
                    if (k < GenreLst.Count())
                    {
                        table1.InsertPageBreakAfterSelf();
                    }
                }
                doc.SaveAs(outfilepath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return File(outfilepath, "application/ms-word", "shourumingxibiao.docx");
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
