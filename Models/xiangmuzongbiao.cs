using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Web.Mvc;

namespace yusuanxiangmu.Models
{
    //public enum shenhezhuangtai
    //{
    //    未审核, 未通过, 已通过
    //}

    public class xiangmuzongbiao
    {
        //预算支出汇总表的字段
        
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ID { get; set; }


        [StringLength(100, ErrorMessage = "不能超过50个汉字。")]
        [Display(Name = "上财项目号")]
        public string shangcaixiangmuhao { get; set; }


        [StringLength(500, ErrorMessage = "不能超过250个汉字。")]
        [Display(Name = "项目名称")]
        public string mingcheng { get; set; }


        // 教学小计、科研小计、行政小计、后勤小计、离退休小计
        [StringLength(50, ErrorMessage = "不能超过25个汉字。")]
        [Display(Name = "项目类别")]
        public string xiangmuleibie { get; set; }

        // 网上指标、基本户、校内追加 或称为 资金来源
        [StringLength(50, ErrorMessage = "不能超过25个汉字。")]
        [Display(Name = "资金来源")]
        public string zijinlaiyuan { get; set; }

        // 学校、个人
        [StringLength(50, ErrorMessage = "不能超过25个汉字。")]
        [Display(Name = "项目分类")]
        public string xiangmufenlei { get; set; }                                //这里的zijingguishu中的jing 多了一个g
                                                                                

        [StringLength(100, ErrorMessage = "不能超过50个汉字。")]
        [Display(Name = "列入栏目")]
        public string lierulanmu { get; set; }


        [StringLength(100, ErrorMessage = "不能超过50个汉字。")]
        [Display(Name = "项目来源")]                                                 //这里是对外的文件编号，如琼财教
        public string xiangmulaiyuan { get; set; }


        [StringLength(100, ErrorMessage = "不能超过50个汉字。")]
        [Display(Name = "编号")]                                                 //这里是对内的文件编号，如  海师财教
        public string bianhao { get; set; }


        [StringLength(50, ErrorMessage = "不能超过25个汉字。")]
        [Display(Name = "执行单位")]
        public string zhixingdanwei { get; set; }



        [StringLength(20, ErrorMessage = "不能超过10个汉字。")]
        [Display(Name = "联系人")]
        public string lianxiren { get; set; }


        [Required(ErrorMessage = "必须输入项目金额")]
        [Display(Name = "项目金额（万元）")]
        [DisplayFormat(DataFormatString = "{0:N4}")]
        [Range(typeof(decimal), "0.0000", "99999999.9999", ErrorMessage = "输入金额格式不正确，小数点后保留4位")]
        [RegularExpression(@"^(([0-9]+)|([0-9]+\.[0-9]{1,4}))$", ErrorMessage = "输入金额格式不正确，小数点后保留4位！")]
        public decimal jine { get; set; }       


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "下达日期")]
        public DateTime xiadariqi { get; set; }


        [StringLength(100, ErrorMessage = "不能超过50个汉字。")]
        [Display(Name = "下达情况")]
        public string xiadaqingkuang { get; set; }


        //以下为预算收入表内容

        //资金来源 见前面

        //项目来源 同上面的 文件编号

        [StringLength(50, ErrorMessage = "不能超过25个汉字。")]
        [Display(Name = "资金性质")]
        public string zijinxingzhi { get; set; }

        [StringLength(50, ErrorMessage = "不能超过25个汉字。")]
        [Display(Name = "下拨部门")]
        public string xiabobumen { get; set; }


        //以下内容由系统自动添加

        [UIHint("HiddenInput")]
        [StringLength(20, ErrorMessage = "不能超过10个汉字。")]
        [Display(Name = "录入人")]
        public string lururen { get; set; }


        //[HiddenInput(DisplayValue = false)]

        [UIHint("HiddenInput")]
        // [ScaffoldColumn(false)] 隐藏此项，ScaffoldColumn  表示的是是否采用MVC框架来处理 设置为true表示采用MVC框架来处理，如果设置为false，则该字段不会在View层显示，里面定义的验证也不会生效。
        [StringLength(10, ErrorMessage = "不能超过5个汉字。")]
        [Display(Name = "录入日期")]
        public string lururiqi { get; set; }
        
        [UIHint("HiddenInput")]
        [StringLength(6, ErrorMessage = "不能超过3个汉字。")]
        [Display(Name = "审核状态")]
        public string shenhezhuangtai { get; set; }
        

        [UIHint("HiddenInput")]
        [StringLength(20, ErrorMessage = "不能超过10个汉字。")]
        [Display(Name = "审核人")]
        public string shenheren { get; set; }
        
        //[HiddenInput(DisplayValue = false)]

        [UIHint("HiddenInput")]
        // [ScaffoldColumn(false)] 隐藏此项，ScaffoldColumn  表示的是是否采用MVC框架来处理 设置为true表示采用MVC框架来处理，如果设置为false，则该字段不会在View层显示，里面定义的验证也不会生效。
        [StringLength(10, ErrorMessage = "不能超过5个汉字。")]
        [Display(Name = "审核日期")]
        public string shenheriqi { get; set; }

        [StringLength(300, ErrorMessage = "不能超过150个汉字。")]
        [Display(Name = "备注")]
        public string beizhu { get; set; }

        // 年度
        [Display(Name = "年度")]
        public int niandu { get; set; }


        //[StringLength(30, ErrorMessage = "不能超过10个汉字。")]                                       //新增加一个  编号 ，和之前的文件编号不同
        //[Display(Name = "编号")]
        //public int bianhao { get; set; }

        [StringLength(30, ErrorMessage = "不能超过10个汉字。")]                                                
        [Display(Name = "校长办公会审议结果")]
        public string xiaozhanghui { get; set; }


        [StringLength(30, ErrorMessage = "不能超过10个汉字。")]
        [Display(Name = "党委会审议结果")]
        public string dangweihui { get; set; }
        

        [StringLength(300, ErrorMessage = "不能超过140个汉字。")]
        [Display(Name = "审批流程")]
        public string shenpiliucheng { get; set; }

     }


}