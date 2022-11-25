using System.ComponentModel.DataAnnotations;


namespace yusuanxiangmu.Models
{
    public class zhixingdanwei
    {

        //预算支出汇总表的字段
        [Key]       
        [StringLength(100, ErrorMessage = "不能超过50个汉字。")]
        [Display(Name = "执行单位名称")]
        public string mingcheng { get; set; }
                
    }
}