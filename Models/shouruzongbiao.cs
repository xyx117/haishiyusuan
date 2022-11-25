using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace yusuanxiangmu.Models
{
    public class shouruzongbiao
    {
        [Display(Name = "项目ID")]
        public int ID { get; set; }

        [Display(Name = "项目名称")]
        public string mingcheng { get; set; }
        
        //对外的编号
        [Display(Name = "项目来源")]
        public string xiangmulaiyuan { get; set; }


        [Display(Name = "上财项目号")]
        public string shancaixiangmuhao { get; set; }     


        [Display(Name = "项目类别")]
        public string xiangmuleibie { get; set; }


        [Display(Name = "资金来源")]
        public string zijinlaiyuan { get; set; }


        [Display(Name = "资金性质")]
        public string zijinxingzhi { get; set; }

        [Display(Name = "执行单位")]
        public string zhixingdanwei { get; set; }


         [Display(Name = "下拨部门")]
        public string xiabobumen { get; set; }


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
         [Display(Name = "下达日期")]
        public DateTime xiadariqi { get; set; }


        [Required(ErrorMessage = "必须输入项目金额")]
        [Display(Name = "项目金额（万元）")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        [Range(typeof(decimal), "0.0000", "99999999.9999", ErrorMessage = "输入金额格式不正确，小数点后保留4位")]
        [RegularExpression(@"^(([0-9]+)|([0-9]+\.[0-9]{1,4}))$", ErrorMessage = "输入金额格式不正确，小数点后保留4位！")]
        public decimal jine { get; set; }


         [Display(Name = "年度")]
         public int niandu { get; set; }


        [Display(Name = "项目类别")]
        public string xiangmufenlei { get; set; }

    }
}