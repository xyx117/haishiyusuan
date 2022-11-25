using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace yusuanxiangmu.Models
{
    public class zhichuzongbiao
    {

            
        [Display(Name = "项目ID")]
        public int ID { get; set; }

        [Display(Name = "项目名称")]
        public string mingcheng { get; set; }

         [Display(Name = "上财项目号")]
        public string shangcaixiangmuhao{ get; set; }

        [Display(Name = "列入栏目")]
        public string lierulanmu{ get; set; }

        [Display(Name = "项目来源")]
        public string xiangmulaiyuan{ get; set; }

        [Display(Name = "项目类别")]
        public string xiangmuleibie { get; set; }

        [Display(Name = "执行单位")]
        public string zhixingdanwei{ get; set; }

        [Display(Name = "联系人")]
        public string lianxiren{ get; set; }

        [Required(ErrorMessage = "必须输入项目金额")]
        [Display(Name = "项目金额（万元）")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        [Range(typeof(decimal), "0.0000", "99999999.9999", ErrorMessage = "输入金额格式不正确，小数点后保留4位")]
        [RegularExpression(@"^(([0-9]+)|([0-9]+\.[0-9]{1,4}))$", ErrorMessage = "输入金额格式不正确，小数点后保留4位！")]
      
        public decimal jine{ get; set; }

        [Display(Name = "项目分类")]                    //这是新加上的
        public string xiangmufenlei { get; set; }

        [Display(Name = "资金来源")]                    //这是新加上的
        public string zijinlaiyuan { get; set; }

        [Display(Name = "审核人")]                    //这是新加上的
        public string shenheren { get; set; }

        [Display(Name = "审核状态")]                    //这是新加上的
        public string shenhezhuangtai { get; set; }

        [Display(Name = "下拨部门")]                    //这是新加上的
        public string xiabobumen { get; set; }

        [Display(Name = "年度")]
        public int niandu { get; set; }
        
        [Display(Name = "校长办公会审议结果")]
        public string xiaozhanghui { get; set; }
        
        [Display(Name = "党委会审议结果")]
        public string dangweihui { get; set; }

        [Display(Name = "审批流程")]                    //这是新加上的
        public string shenpiliucheng { get; set; }

        }
}