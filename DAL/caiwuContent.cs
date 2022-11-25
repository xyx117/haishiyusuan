using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;
using System.Linq;
using System.Web;
using yusuanxiangmu.Models;


namespace yusuanxiangmu.DAL
{
    public class caiwuContent:DbContext
    {
        public DbSet<xiangmuzongbiao> xiangmuzongbiaos{get; set; }
        public DbSet<zhixingdanwei> zhixingdanweis { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();

            modelBuilder.Entity<xiangmuzongbiao>().Property(xiangmuzongbiao => xiangmuzongbiao.jine).HasPrecision(12, 4); 

            modelBuilder.Entity<xiangmuzongbiao>().MapToStoredProcedures();

            modelBuilder.Entity<zhixingdanwei>().MapToStoredProcedures();
        }     
    }
}